*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${False}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF
Library             RPA.Robocorp.Vault


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Log in
    Download the Excel file
    Fill the form using the data from the Excel file
    #Fill and submit the form for one person    $sales_rep
    Collect the results
    Export the table as pdf
    [Teardown]    Log out and close browser


*** Keywords ***
Open the intranet website
    Open Available Browser    https://robotsparebinindustries.com/    maximized=${True}

Log in
    ${secret}=    Get Secret    robotsparebin
    Input Text    username    ${secret}[username]
    Input Password    password    ${secret}[password]
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download the Excel file
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=${True}

Fill and submit the form for one person
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[First Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Input Text    salesresult    ${sales_rep}[Sales]
    #Select From List By Value is used when there is a dropdown
    Select From List By Value    id:salestarget    ${sales_rep}[Sales Target]
    Click Button    Submit

Fill the form using the data from the Excel file
    Open Workbook    SalesData.xlsx
    #Saves the table in sales_reps variable
    ${sales_reps}=    Read Worksheet As Table    header=${True}
    Close Workbook
    #Loops over the sales_reps; loop like python
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and submit the form for one person    ${sales_rep}
    END

Collect the results
    #Takes screenshot of the css div locator which has a class 'sales-summary'. Here the
    #${OUTPUT_DIR} is a robot framework runtime varialbe that represents output directory.
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Export the table as pdf
    Wait Until Element Is Visible    id:sales-results
    ${sales_result_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_result_html}    ${OUTPUT_DIR}${/}sales_resulf.pdf

Log out and close browser
    Click Button    Log out
    Close Browser
