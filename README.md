1. Call Monitoring ITSD Audit Template

- The primary sheets within this template are labeled "Audit Form" and "Audit Viewer."
- Users will input data into the form; certain fields will auto-populate upon selection of an analyst.
- Percentages will be automatically calculated once checkboxes are marked.
- Upon correct input of all data, users can save the information, which will then be transmitted to the "Audit Viewer."
- Subsequently, saved data within the Audit Viewer can be emailed via a designated button.
- Utilizing VBA, an array facilitates the linking of cells between Sheet1 and Sheet2.
- Further insights into functionality can be gleaned by reviewing the accompanying VBA code.


2. Ticket Hygiene

  Button: 'Add Record'
- By selecting the 'Add Record' button, users initiate the addition of a new entry.
- This action triggers a user interface prompt.
- Analysts will conduct random ticket audits, adhering to specified criteria outlined within the form.
- It is imperative that all fields are completed accurately to avoid potential errors.
- Should any fields remain unfilled, a notification will alert the user.
- Upon submission, data is saved directly to the Excel file.

  Button: 'Clear'
- Designed to reset all data entries, the 'Clear' button enables the reuse of the Excel file.

  Button: 'Email'
- With a simple click, users can effortlessly send the file via email utilizing the designated button.
