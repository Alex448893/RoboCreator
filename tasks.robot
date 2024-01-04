*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF
Library             RPA.Robocorp.WorkItems
Library             RPA.Tables
Library             RPA.Desktop
Library             RPA.Archive
Library             OperatingSystem
Library             RPA.Email.Exchange


*** Tasks ***
Build the robot and export it as a PDF
    Open the internet website and Click the popup
    Download the Excel file
    Fill the form using the Excel file
    Zip the PDFs


*** Keywords ***
Open the internet website and Click the popup
    Open Available Browser    https://robotsparebinindustries.com/#/robot-order
    log    ${OUTPUT_DIR}
    Create Directory    receipts

Download the Excel file
    Download    https://robotsparebinindustries.com/orders.csv    overwrite=True

Fill and submit the form for one robot
    [Arguments]    ${order}
    Click Button    OK

    # Maximize Browser Window
    Set Window Size    1920    1080
    Select From List By Index    head    ${order}[Head]
    Click Button    id-body-${order}[Body]
    Input Text    class:form-control    ${order}[Legs]
    Input Text    address    ${order}[Address]
    Click Button When Visible    preview

    # Click Element    preview Order number,Head,Body,Legs,Address

Fill the form using the Excel file
    # Open Workbook    orders.csv
    ${orders}=    Read table from CSV    orders.csv    header=True
    Close Workbook
    FOR    ${order}    IN    @{orders}
        Fill and submit the form for one robot    ${order}
        Order the robot and document the receipt    ${order}
    END

Order the robot and document the receipt
    [Arguments]    ${order}

    Click Button    order
    Store the receipt as a PDF file    ${order}

    # TRY
    #    Click Button    order-another
    # EXCEPT    Button with locator 'order-another' not found.
    #    ${error}=    Set Variable    True
    #    WHILE    ${error} == True
    #    Click Button    order
    #    ${error}=    Is Element Visible    class=alert-danger
    #    END
    #    Click Button    order-another
    # EXCEPT    Button with locator 'OK' not found.
    #    Log    message
    # EXCEPT
    #    Click Button    order-another
    # END
    Click Button    order-another

Store the receipt as a PDF file
    [Arguments]    ${order}

    TRY
        ${receipt_html}=    Get Element Attribute    id:receipt    outerHTML
    EXCEPT    Element with locator 'id:receipt' not found.
        ${error}=    Set Variable    True
        WHILE    ${error} == True
            Click Button    order
            ${error}=    Is Element Visible    class=alert-danger
        END
        ${receipt_html}=    Get Element Attribute    id:receipt    outerHTML
    END
    # ${Screenshot}=
    ${Screenshot}=    RPA.Browser.Selenium.Screenshot    robot-preview-image
    ${receipt_html}=    Get Element Attribute    id:receipt    outerHTML
    Html To Pdf    ${receipt_html}    ${OUTPUT_DIR}${/}${order}[Order number].pdf
    ${files}=    Create List    ${OUTPUT_DIR}${/}${order}[Order number].pdf    ${Screenshot}
    # Add Files To Pdf    ${Screenshot}    ${OUTPUT_DIR}${/}${order}[Order number].pdf    True
    Add Files To Pdf    ${files}    ${OUTPUT_DIR}${/}${order}[Order number].pdf
    Move File    ${OUTPUT_DIR}${/}${order}[Order number].pdf    receipts

Zip the Pdfs
    # Move logs to output
    # Zip receipts
    Archive Folder With Zip    receipts    receipts.zip    TRUE
