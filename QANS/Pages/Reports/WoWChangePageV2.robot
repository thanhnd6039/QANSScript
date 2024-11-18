*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${wowChangeResultFilePath}        C:\\RobotFramework\\Results\\WoWChangeReportResult.xlsx
${wowChangeFilePath}              C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${wowChangeOnVDCFilePath}         C:\\RobotFramework\\Downloads\\Wow Change [Current Week] On VDC.xlsx
${SGFilePath}                     C:\\RobotFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx

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

Check Data For The OEM East Table
    [Arguments]     ${wowChangeFilePath}  ${SGFilePath}   ${posOfColOnWoWChange}    ${posOfColOnSG}     ${nameOfCol}
    ${result}   Set Variable    ${True}
    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}           doc_id=WoWChange

    File Should Exist      path=${SGFilePath}
    Open Excel Document    filename=${SGFilePath}    doc_id=SG
    Switch Current Excel Document     doc_id=SG
    ${numOfRowsOnSG}    Get Number Of Rows In Excel    ${SGFilePath}
    Switch Current Excel Document    doc_id=WoWChange

#   Verify the data for each OEM Group
    FOR    ${rowIndexOnWoWChange}    IN RANGE    2    7
        Switch Current Excel Document    doc_id=WoWChange
        ${oemGroupColOnWoWChange}      Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1        
        ${dataColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
        Switch Current Excel Document     doc_id=SG
        FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
            ${oemGroupColOnSG}       Read Excel Cell    row_num=${rowIndexOnSG}    col_num=2
            IF    '${oemGroupColOnSG}' == 'None'
                 Continue For Loop
            END            
            ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfColOnSG}
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
        ${dataColOnSG}              Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfColOnSG}
        IF    '${mainSalesRepColOnSG}' == 'None'
            Continue For Loop
        END
        IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
             IF    '${dataColOnSG}' == 'None'
                  ${dataColOnSG}    Set Variable    0
             END
             ${totalOnSG}     Evaluate    ${totalOnSG}+${dataColOnSG}
        END
    END
    ${totalOnSG}   Evaluate  "%.2f" % ${totalOnSG}

    Switch Current Excel Document    doc_id=WoWChange
    ${totalOnWoWchange}   Read Excel Cell    row_num=8    col_num=${posOfColOnWoWChange}
    IF    ${totalOnWoWchange} != ${totalOnSG}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total    ${totalOnWoWchange}    ${totalOnSG}
    END

 #  Verify the OTHERS data
    Switch Current Excel Document    doc_id=SG
    ${othersOnSG}  Set Variable    0
    FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
        ${oemGroupColOnSG}          Read Excel Cell    row_num=${rowIndexOnSG}    col_num=2
        ${mainSalesRepColOnSG}      Read Excel Cell    row_num=${rowIndexOnSG}    col_num=3
        ${dataColOnSG}              Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfColOnSG}
        IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
             IF    '${oemGroupColOnSG}' not in ${listOfOEMGroupShownInOEMEastTable}
                  IF    '${dataColOnSG}' == 'None'
                       ${dataColOnSG}   Set Variable    0
                  END
                  ${othersOnSG}     Evaluate    ${othersOnSG}+${dataColOnSG}
             END
        END
    END
    ${othersOnSG}   Evaluate  "%.2f" % ${othersOnSG}
    Switch Current Excel Document    doc_id=WoWChange
    ${othersOnWoWChange}   Read Excel Cell    row_num=7    col_num=${posOfColOnWoWChange}
    IF    ${othersOnWoWChange} != ${othersOnSG}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East OTHERS    ${othersOnWoWChange}    ${othersOnSG}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the OEM East table between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents
    