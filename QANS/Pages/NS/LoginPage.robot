*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${txtEmail}        //*[@id='email']
${txtPassword}     //*[@id='password']
${btnLogin}        //*[@id='login-submit']
${txtLoginTitle}    //*[@id='uif49']
${txtAccountTitle}  //*[@class='uir-roleswitch-section-title']
${txtVerificationCode}    //*[@id='uif56_input']
${chkTrustThisDeviceFor30Days}    //*[@id='uif70_check']
${btnSubmit}    //label[normalize-space()='Submit']
${imgSANDBOXIcon}   //*[@aria-label='sandbox']
${imgLogoIcon}      //img[@id='uif40']

*** Keywords ***
Login To NS With Account
    [Arguments]     ${account}
    Login To NS
    Should See Account Title    Choose account
    Choose Account    ${account}
    IF    '${account}' == 'PRODUCTION'
        Should See Login Title    Logging in to Virtium
    ELSE IF     '${account}' == 'SANDBOX1'
        Should See Login Title    Logging in to Virtium_SB1
    END
    Input Verification Code And Click Submit
    IF    '${account}' == 'PRODUCTION'
        Should See Account Is PRODUCTION
    ELSE IF     '${account}' == 'SANDBOX1'
        Should See Account Is SANDBOX
    END

Login To NS
    ${configFileObject}     Load Json From File    ${CONFIG_DIR}\\NSConfig.json
    ${url}                  Get Value From Json    ${configFileObject}    $.url
    ${url}               Set Variable       ${url[0]}
    ${email}             Get Value From Json    ${configFileObject}    $.accounts[0].email
    ${email}             Set Variable       ${email[0]}
    ${pass}              Get Value From Json    ${configFileObject}    $.accounts[0].password
    ${pass}              Set Variable       ${pass[0]}
    Go To    ${url}
    Wait Until Element Is Enabled    ${txtEmail}       ${TIMEOUT}
    Input Text    ${txtEmail}    ${email}
    Wait Until Element Is Enabled    ${txtPassword}    ${TIMEOUT}
    Input Text    ${txtPassword}    ${pass}
    Wait Until Element Is Enabled    ${btnLogin}       ${TIMEOUT}
    Click Element    ${btnLogin}

Choose Account
    [Arguments]     ${account}
    @{listOfAccountElements}    Get WebElements    //*[@class='uir-roleswitch-table']/tbody/tr
    ${numOfAccounts}     Get Length    ${listOfAccountElements}

    FOR    ${accountIndex}    IN RANGE    2    ${numOfAccounts}+1
        ${company}      Get Text    //*[@class='uir-roleswitch-table']/tbody/tr[${accountIndex}]/td[1]
        ${accountType}  Get Text    //*[@class='uir-roleswitch-table']/tbody/tr[${accountIndex}]/td[2]
        IF    '${account}' == 'PRODUCTION'
            IF    '${company}' == 'Virtium' and '${accountType}' == 'PRODUCTION'
                 Wait Until Element Is Visible    //*[@class='uir-roleswitch-table']/tbody/tr[${accountIndex}]/td[3]/a     ${TIMEOUT}
                 Click Element    //*[@class='uir-roleswitch-table']/tbody/tr[${accountIndex}]/td[3]/a
                 Exit For Loop
            END
        ELSE IF     '${account}' == 'SANDBOX1'
            IF    '${company}' == 'Virtium_SB1' and '${accountType}' == 'SANDBOX'
                Wait Until Element Is Visible    //*[@class='uir-roleswitch-table']/tbody/tr[${accountIndex}]/td[3]/a      ${TIMEOUT}
                Click Element    //*[@class='uir-roleswitch-table']/tbody/tr[${accountIndex}]/td[3]/a
                Exit For Loop
            END
        ELSE
             Fail   The Account parameter ${account} is invalid
        END
    END

Should See Login Title
    [Arguments]     ${title}
    Wait Until Element Is Visible    ${txtLoginTitle}       ${TIMEOUT}
    Element Text Should Be           ${txtLoginTitle}       ${title}

Should See Account Title
    [Arguments]     ${title}
    Wait Until Element Is Visible    ${txtAccountTitle}       ${TIMEOUT}
    Element Text Should Be    ${txtAccountTitle}    ${title}

Input Verification Code And Click Submit
    ${configFileObject}     Load Json From File    ${CONFIG_DIR}\\NSConfig.json
    ${key}      Get Value From Json    ${configFileObject}    $.key
    ${key}      Set Variable    ${key[0]}
    ${otp}      Generate Otp    ${key}
    Wait Until Element Is Enabled    ${txtVerificationCode}     ${TIMEOUT}
    Input Text    ${txtVerificationCode}    ${otp}
    Wait Until Element Is Enabled    ${chkTrustThisDeviceFor30Days}     ${TIMEOUT}
    Click Element    ${chkTrustThisDeviceFor30Days}
    Wait Until Element Is Enabled    ${btnSubmit}   ${TIMEOUT}
    Click Element    ${btnSubmit}

Should See Account Is SANDBOX
    Wait Until Element Is Visible    ${imgSANDBOXIcon}      ${TIMEOUT}

Should See Account Is PRODUCTION
    Wait Until Element Is Visible    ${imgLogoIcon}     ${TIMEOUT}
    
