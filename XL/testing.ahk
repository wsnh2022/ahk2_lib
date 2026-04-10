; XL.ahk Testing Script
; Comprehensive examples demonstrating the XL.ahk library functionality
#Include "C:\Users\The_Thinker\Documents\ahk2_lib\XL\XL.ahk"

; Example 1: Basic Excel file creation with formatted data
Example1_BasicData()

; Example 2: Rich text and font formatting
Example2_RichTextFormatting()

; Example 3: Formulas and calculations
Example3_FormulasAndCalculations()

; Example 4: Date and number formatting
Example4_DateTimeAndNumbers()

; Example 5: Reading existing Excel file
Example5_ReadExistingFile()

; Example 6: Multiple sheets with different data types
Example6_MultipleSheets()

; Example 7: Cell styling and borders
Example7_CellStylingAndBorders()

; Example 8: Working with ranges and merged cells
Example8_RangesAndMergedCells()

MsgBox("All examples completed! Check the generated Excel files.")

; Basic data entry and file creation
Example1_BasicData() {
    ; Create new Excel workbook
    book := XL.New('xlsx')
    sheet := book.addSheet('Basic Data')
    
    ; Add headers
    sheet['A1'] := "Name"
    sheet['B1'] := "Age"
    sheet['C1'] := "City"
    sheet['D1'] := "Salary"
    
    ; Add sample data
    sheet['A2'] := "John Doe"
    sheet['B2'] := 30
    sheet['C2'] := "New York"
    sheet['D2'] := 50000
    
    sheet['A3'] := "Jane Smith"
    sheet['B3'] := 25
    sheet['C3'] := "Los Angeles"
    sheet['D3'] := 45000
    
    sheet['A4'] := "Bob Johnson"
    sheet['B4'] := 35
    sheet['C4'] := "Chicago"
    sheet['D4'] := 55000
    
    ; Save the workbook
    book.save('Example1_BasicData.xlsx')
    book := ''
}

; Rich text formatting with multiple fonts and colors
Example2_RichTextFormatting() {
    book := XL.New('xlsx')
    sheet := book.addSheet('Rich Text')
    
    ; Create rich string with different formatting
    rs := book.addRichString()
    
    ; Different font styles for "E=mc²"
    font1 := rs.addFont()
    font1.setColor(10)  ; Red
    font1.setSize(24)
    font1.setBold(true)
    
    font2 := rs.addFont()
    font2.setSize(24)
    font2.setItalic(true)
    
    font3 := rs.addFont()
    font3.setColor(12)  ; Blue
    font3.setSize(24)
    
    font4 := rs.addFont()
    font4.setColor(17)  ; Green
    font4.setSize(24)
    
    font5 := rs.addFont()
    font5.setScript(1)  ; Superscript
    font5.setSize(16)
    
    ; Build the rich text
    rs.addText('E', font1)
    rs.addText('=', font2)
    rs.addText('m', font3)
    rs.addText('c', font4)
    rs.addText('²', font5)
    
    sheet['A1'] := rs
    
    ; Add some regular formatted text
    headerFormat := book.addFormat()
    headerFormat.font().setBold(true)
    headerFormat.font().setSize(14)
    headerFormat.setAlignH(2)  ; Center align
    
    sheet['A3'] := {value: "Rich Text Examples", format: headerFormat}
    
    book.save('Example2_RichText.xlsx')
    book := ''
}

; Mathematical formulas and calculations
Example3_FormulasAndCalculations() {
    book := XL.New('xlsx')
    sheet := book.addSheet('Formulas')
    
    ; Set calculation mode to manual for better control
    book.setCalcMode(0)
    
    ; Add some numbers for calculations
    sheet['A1'] := "Value A"
    sheet['B1'] := "Value B"
    sheet['C1'] := "Sum"
    sheet['D1'] := "Product"
    sheet['E1'] := "Average"
    
    sheet['A2'] := 10
    sheet['B2'] := 20
    sheet['A3'] := 15
    sheet['B3'] := 25
    sheet['A4'] := 30
    sheet['B4'] := 40
    
    ; Add formulas
    sheet['C2'] := {expr: 'A2+B2'}
    sheet['C3'] := {expr: 'A3+B3'}
    sheet['C4'] := {expr: 'A4+B4'}
    
    sheet['D2'] := {expr: 'A2*B2'}
    sheet['D3'] := {expr: 'A3*B3'}
    sheet['D4'] := {expr: 'A4*B4'}
    
    sheet['E2'] := {expr: '(A2+B2)/2'}
    sheet['E3'] := {expr: '(A3+B3)/2'}
    sheet['E4'] := {expr: '(A4+B4)/2'}
    
    ; Add totals
    sheet['A6'] := "TOTALS:"
    sheet['C6'] := {expr: 'SUM(C2:C4)'}
    sheet['D6'] := {expr: 'SUM(D2:D4)'}
    sheet['E6'] := {expr: 'AVERAGE(E2:E4)'}
    
    book.save('Example3_Formulas.xlsx')
    book := ''
}

; Date/time formatting and number formats
Example4_DateTimeAndNumbers() {
    book := XL.New('xlsx')
    sheet := book.addSheet('Dates & Numbers')
    
    ; Headers
    sheet['A1'] := "Date"
    sheet['B1'] := "Time"
    sheet['C1'] := "Currency"
    sheet['D1'] := "Percentage"
    sheet['E1'] := "Scientific"
    
    ; Create different number formats
    dateFormat := book.addFormat()
    dateFormat.setNumFormat(14)  ; Date format
    
    timeFormat := book.addFormat()
    timeFormat.setNumFormat(21)  ; Time format
    
    currencyFormat := book.addFormat()
    currencyFormat.setNumFormat(5)  ; Currency format
    
    percentFormat := book.addFormat()
    percentFormat.setNumFormat(10)  ; Percentage format
    
    scientificFormat := book.addFormat()
    scientificFormat.setNumFormat(11)  ; Scientific format
    
    ; Add formatted data
    sheet['A2'] := {value: book.datePack(2024, 1, 15, 0, 0, 0), format: dateFormat}
    sheet['A3'] := {value: book.datePack(2024, 6, 20, 0, 0, 0), format: dateFormat}
    sheet['A4'] := {value: book.datePack(2024, 12, 25, 0, 0, 0), format: dateFormat}
    
    sheet['B2'] := {value: book.datePack(2024, 1, 1, 9, 30, 0), format: timeFormat}
    sheet['B3'] := {value: book.datePack(2024, 1, 1, 14, 45, 30), format: timeFormat}
    sheet['B4'] := {value: book.datePack(2024, 1, 1, 18, 15, 45), format: timeFormat}
    
    sheet['C2'] := {value: 1234.56, format: currencyFormat}
    sheet['C3'] := {value: 789.12, format: currencyFormat}
    sheet['C4'] := {value: 2500.99, format: currencyFormat}
    
    sheet['D2'] := {value: 0.85, format: percentFormat}
    sheet['D3'] := {value: 1.25, format: percentFormat}
    sheet['D4'] := {value: 0.05, format: percentFormat}
    
    sheet['E2'] := {value: 123456789, format: scientificFormat}
    sheet['E3'] := {value: 0.000000123, format: scientificFormat}
    sheet['E4'] := {value: 987654.321, format: scientificFormat}
    
    ; Set column widths
    sheet.setCol(0, 0, 12)  ; Column A
    sheet.setCol(1, 1, 10)  ; Column B
    sheet.setCol(2, 2, 12)  ; Column C
    sheet.setCol(3, 3, 12)  ; Column D
    sheet.setCol(4, 4, 15)  ; Column E
    
    book.save('Example4_DateTime.xlsx')
    book := ''
}

; Reading an existing Excel file
Example5_ReadExistingFile() {
    ; First create a file to read
    book := XL.New('xlsx')
    sheet := book.addSheet('Sample Data')
    
    sheet['A1'] := "Product"
    sheet['B1'] := "Quantity"
    sheet['C1'] := "Price"
    
    sheet['A2'] := "Widget A"
    sheet['B2'] := 100
    sheet['C2'] := 19.99
    
    sheet['A3'] := "Widget B"
    sheet['B3'] := 150
    sheet['C3'] := 24.99
    
    book.save('SampleData.xlsx')
    book := ''
    
    ; Now read the file back
    book := XL.Load('SampleData.xlsx')
    sheet := book.getSheet(0)
    
    ; Read and display data
    productName := sheet['A2'].value
    quantity := sheet['B2'].value
    price := sheet['C2'].value
    
    MsgBox("Read from Excel:`nProduct: " . productName . "`nQuantity: " . quantity . "`nPrice: $" . price)
    
    book := ''
}

; Multiple sheets with different data types
Example6_MultipleSheets() {
    book := XL.New('xlsx')
    
    ; Sheet 1: Employee Data
    employeeSheet := book.addSheet('Employees')
    employeeSheet['A1'] := "ID"
    employeeSheet['B1'] := "Name"
    employeeSheet['C1'] := "Department"
    employeeSheet['D1'] := "Hire Date"
    
    employeeSheet['A2'] := 1001
    employeeSheet['B2'] := "Alice Johnson"
    employeeSheet['C2'] := "Engineering"
    employeeSheet['D2'] := "2022-01-15"
    
    employeeSheet['A3'] := 1002
    employeeSheet['B3'] := "Bob Smith"
    employeeSheet['C3'] := "Marketing"
    employeeSheet['D3'] := "2021-08-20"
    
    ; Sheet 2: Sales Data
    salesSheet := book.addSheet('Sales')
    salesSheet['A1'] := "Month"
    salesSheet['B1'] := "Revenue"
    salesSheet['C1'] := "Target"
    salesSheet['D1'] := "Achievement %"
    
    salesSheet['A2'] := "January"
    salesSheet['B2'] := 50000
    salesSheet['C2'] := 45000
    salesSheet['D2'] := {expr: 'B2/C2'}
    
    salesSheet['A3'] := "February"
    salesSheet['B3'] := 55000
    salesSheet['C3'] := 50000
    salesSheet['D3'] := {expr: 'B3/C3'}
    
    ; Sheet 3: Inventory
    inventorySheet := book.addSheet('Inventory')
    inventorySheet['A1'] := "Item"
    inventorySheet['B1'] := "In Stock"
    inventorySheet['C1'] := "Min Level"
    inventorySheet['D1'] := "Reorder Needed"
    
    inventorySheet['A2'] := "Laptop"
    inventorySheet['B2'] := 25
    inventorySheet['C2'] := 10
    inventorySheet['D2'] := {expr: 'IF(B2<C2,"YES","NO")'}
    
    inventorySheet['A3'] := "Monitor"
    inventorySheet['B3'] := 5
    inventorySheet['C3'] := 15
    inventorySheet['D3'] := {expr: 'IF(B3<C3,"YES","NO")'}
    
    book.save('Example6_MultipleSheets.xlsx')
    book := ''
}

; Cell styling, borders, and colors
Example7_CellStylingAndBorders() {
    book := XL.New('xlsx')
    sheet := book.addSheet('Styled Cells')
    
    ; Create various formats
    headerFormat := book.addFormat()
    headerFormat.font().setBold(true)
    headerFormat.font().setSize(14)
    headerFormat.font().setColor(9)  ; White
    headerFormat.setFillPattern(1)   ; Solid fill
    headerFormat.setPatternForegroundColor(12)  ; Blue background
    headerFormat.setAlignH(2)        ; Center align
    headerFormat.setBorder(1)        ; Thin border
    
    dataFormat := book.addFormat()
    dataFormat.setBorder(1)
    dataFormat.setAlignH(1)          ; Left align
    
    highlightFormat := book.addFormat()
    highlightFormat.setFillPattern(1)
    highlightFormat.setPatternForegroundColor(13)  ; Yellow background
    highlightFormat.setBorder(1)
    
    ; Apply formatting
    sheet['A1'] := {value: "Styled Header", format: headerFormat}
    sheet['B1'] := {value: "Another Header", format: headerFormat}
    sheet['C1'] := {value: "Third Header", format: headerFormat}
    
    sheet['A2'] := {value: "Normal Data", format: dataFormat}
    sheet['B2'] := {value: "More Data", format: dataFormat}
    sheet['C2'] := {value: "Even More", format: dataFormat}
    
    sheet['A3'] := {value: "Highlighted", format: highlightFormat}
    sheet['B3'] := {value: "Important!", format: highlightFormat}
    sheet['C3'] := {value: "Notice Me", format: highlightFormat}
    
    ; Set column widths
    sheet.setCol(0, 2, 15)
    
    book.save('Example7_Styling.xlsx')
    book := ''
}

; Working with cell ranges and merged cells
Example8_RangesAndMergedCells() {
    book := XL.New('xlsx')
    sheet := book.addSheet('Ranges & Merges')
    
    ; Create title with merged cells
    titleFormat := book.addFormat()
    titleFormat.font().setBold(true)
    titleFormat.font().setSize(16)
    titleFormat.setAlignH(2)  ; Center align
    titleFormat.setAlignV(1)  ; Middle align
    
    sheet['A1'] := {value: "QUARTERLY SALES REPORT", format: titleFormat}
    sheet.setMerge(0, 0, 0, 4)  ; Merge A1:E1
    
    ; Create section headers
    sectionFormat := book.addFormat()
    sectionFormat.font().setBold(true)
    sectionFormat.setFillPattern(1)
    sectionFormat.setPatternForegroundColor(15)  ; Gray background
    
    sheet['A3'] := {value: "Q1 Results", format: sectionFormat}
    sheet['A4'] := "Product A"
    sheet['A5'] := "Product B"
    sheet['A6'] := "Product C"
    
    sheet['B3'] := {value: "Sales", format: sectionFormat}
    sheet['B4'] := 10000
    sheet['B5'] := 15000
    sheet['B6'] := 8000
    
    sheet['C3'] := {value: "Target", format: sectionFormat}
    sheet['C4'] := 12000
    sheet['C5'] := 14000
    sheet['C6'] := 9000
    
    sheet['D3'] := {value: "Variance", format: sectionFormat}
    sheet['D4'] := {expr: 'B4-C4'}
    sheet['D5'] := {expr: 'B5-C5'}
    sheet['D6'] := {expr: 'B6-C6'}
    
    sheet['E3'] := {value: "% Achievement", format: sectionFormat}
    sheet['E4'] := {expr: 'B4/C4'}
    sheet['E5'] := {expr: 'B5/C5'}
    sheet['E6'] := {expr: 'B6/C6'}
    
    ; Total row
    totalFormat := book.addFormat()
    totalFormat.font().setBold(true)
    totalFormat.setBorderTop(2)  ; Thick top border
    
    sheet['A8'] := {value: "TOTAL", format: totalFormat}
    sheet['B8'] := {expr: 'SUM(B4:B6)', format: totalFormat}
    sheet['C8'] := {expr: 'SUM(C4:C6)', format: totalFormat}
    sheet['D8'] := {expr: 'SUM(D4:D6)', format: totalFormat}
    sheet['E8'] := {expr: 'AVERAGE(E4:E6)', format: totalFormat}
    
    ; Set column widths
    sheet.setCol(0, 0, 12)  ; Product names
    sheet.setCol(1, 4, 10)  ; Numeric columns
    
    ; Set row height for title
    sheet.setRow(0, 25)
    
    book.save('Example8_RangesAndMerges.xlsx')
    book := ''
}