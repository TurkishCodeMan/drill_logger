# Project Overview

The Excel Tablet App project aims to facilitate filling rows in large Excel files by allowing each row to be processed through a wide window with a well-designed UI. Instead of directly working on the Excel file, data will be entered using an intermediary tool and then saved back into the Excel file.
We will be using react,shadcn,tailwind,lucid icon

# Core Functionalities
1. The application will allow uploading an Excel file and selecting a specific sheet. No operations will be performed on other sheets.
    1.1. The user can select an Excel file from their device.
    1.2. The list of sheets in the uploaded Excel file will be displayed on the screen.
    1.3. Only one specific sheet can be selected.
    1.4. No operations will be performed on unselected sheets.
    1.5. Supported file formats: .xlsx and .xls.
    1.6. After the file and sheet selection, the user will proceed to the data processing step.
2. Upon opening the selected Excel sheet, the app will extract the mentioned column names and create groups. These groups will contain the column names from the next row as their elements, with each row's name defined by these column headers.
    2.1. The first row of the selected sheet will be treated as column headers.
    2.2. Automatic Column Detection: Column headers (e.g., “Product Name,” “Price,” “Stock”) will be identified automatically.
    2.3. Group Creation: The app will create groups based on these column headers.
    2.4. Subcolumns: The second row will determine the subcolumns under each group.
    2.5. Example:
    Group: Product Information
    Subcolumns: Product Name, Product Code, Price
    Group: Customer Information
    Subcolumns: First Name, Last Name, Phone Number
    2.6. Flexible Structure: Users can manually adjust or merge groups if needed.
3. On the main page, after uploading the Excel file, the data entry process will start from the first row and focus on filling the columns under the specified groups.
    3.1. Data Entry Start: After uploading the Excel file, the data entry process will begin from the first row on the main page.
    3.2. Group-Focused Input: Only the columns under the defined groups will be filled in for each row.
    3.3. Dynamic Data Flow: Once data entry for one row is completed, the next row will be automatically loaded.
    3.4. Data Validation: The system will check for invalid entries (for instance, letters in a numeric field).
    3.5. Mandatory Fields: Certain columns can be marked as required (e.g., “Product Name” or “Date”).
    3.6. Auto-Save: For long data-entry sessions, an automatic save feature can be utilized to prevent data loss.
4. After making updates, clicking the save button will ensure the entire Excel format, including colors, font styles, font sizes, and formatting, is preserved and saved correctly.
    4.1. Save Operation: After data entry is finished, the user can click the “Save” button to update the Excel file.
    4.2. Format Preservation: The following Excel formatting elements will be preserved when saving:
    Colors: Cell background colors.
    Fonts: Font family, size, and weight.
    Cell Dimensions: Row heights and column widths.
    Borders: Cell borders and line thickness.
    Alignment: Text alignment within cells (center, left, right).
    4.3. No Loss of Original Format: The file will be saved without losing any data or design aspects.
    4.4. Rollback Option: If any error occurs during saving, an undo or rollback option can be provided.
5. Each row will be displayed in a clearly readable size, prioritizing interactive usage for better clarity and usability.
    5.1. Large and Readable View: Each row will be shown in a larger format to simplify data input.
    5.2. Interactive Usage: Users can expand or collapse each row as needed.
    5.3. Tablet-Friendly Design: The interface will be optimized for touch input:
    Larger tap areas.
    Bigger text fields.
    Touch-friendly buttons.
    5.4. Error Highlighting: Any invalid data entry will be clearly indicated (e.g., with a red background).
    5.5. Scrolling Form Structure: As the user completes each row, the form will scroll automatically to the next one.

# Doc
 ## Documentation of ExcelJS load excel file and processing.
 CODE EXAMPLE:
 ```
    import * as ExcelJS from 'exceljs';
    import { ExcelConfig, ExcelData, ExcelGroup, ExcelUploadResponse } from '../types/excel.types';

    const DEFAULT_CONFIG: ExcelConfig = {
    sheetName: 'GEOLOGY',
    groupRow: 3,
    columnRow: 4,
    dataStartRow: 5
    };

    export class ExcelService {
    private workbook: ExcelJS.Workbook | null = null;
    private worksheet: ExcelJS.Worksheet | null = null;

    async loadExcelFile(file: File): Promise<ExcelUploadResponse> {
        try {
        this.workbook = new ExcelJS.Workbook();
        const arrayBuffer = await file.arrayBuffer();
        await this.workbook.xlsx.load(arrayBuffer);

        // GEOLOGY sayfasını bul
        this.worksheet = this.workbook.getWorksheet(DEFAULT_CONFIG.sheetName);
        
        if (!this.worksheet) {
            return {
            success: false,
            message: 'GEOLOGY sayfası bulunamadı',
            error: 'Sheet not found'
            };
        }

        const groups = await this.extractGroups();
        const rows = await this.extractRows(groups);

        return {
            success: true,
            message: 'Excel dosyası başarıyla yüklendi',
            data: {
            groups,
            rows
            }
        };
        } catch (error) {
        return {
            success: false,
            message: 'Excel dosyası yüklenirken hata oluştu',
            error: error instanceof Error ? error.message : 'Unknown error'
        };
        }
    }

    private async extractGroups(): Promise<ExcelGroup[]> {
        if (!this.worksheet) throw new Error('Worksheet not loaded');

        const groups: ExcelGroup[] = [];
        const groupRow = this.worksheet.getRow(DEFAULT_CONFIG.groupRow);
        const columnRow = this.worksheet.getRow(DEFAULT_CONFIG.columnRow);
        
        let currentGroup: ExcelGroup | null = null;
        let columnIndex = 1;

        while (columnIndex <= columnRow.cellCount) {
        const groupCell = groupRow.getCell(columnIndex);
        const columnCell = columnRow.getCell(columnIndex);
        
        if (groupCell.value) {
            if (currentGroup) {
            currentGroup.endColumn = columnIndex - 1;
            groups.push(currentGroup);
            }
            
            currentGroup = {
            name: groupCell.value.toString(),
            color: groupCell.fill?.fgColor?.argb || '',
            columns: [],
            startColumn: columnIndex,
            endColumn: columnIndex
            };
        }

        if (currentGroup && columnCell.value) {
            currentGroup.columns.push(columnCell.value.toString());
        }

        columnIndex++;
        }

        if (currentGroup) {
        currentGroup.endColumn = columnIndex - 1;
        groups.push(currentGroup);
        }

        return groups;
    }

    private async extractRows(groups: ExcelGroup[]): Promise<Record<string, any>[]> {
        if (!this.worksheet) throw new Error('Worksheet not loaded');

        const rows: Record<string, any>[] = [];
        let rowIndex = DEFAULT_CONFIG.dataStartRow;

        while (rowIndex <= this.worksheet.rowCount) {
        const excelRow = this.worksheet.getRow(rowIndex);
        const rowData: Record<string, any> = {};

        groups.forEach(group => {
            for (let i = group.startColumn; i <= group.endColumn; i++) {
            const columnName = this.worksheet!.getRow(DEFAULT_CONFIG.columnRow).getCell(i).value?.toString() || '';
            if (columnName) {
                rowData[columnName] = excelRow.getCell(i).value;
            }
            }
        });

        if (Object.keys(rowData).length > 0) {
            rows.push(rowData);
        }

        rowIndex++;
        }

        return rows;
    }
    }   
 ```


# Current file structure
excel-tablet-app
    ├── README.md
    ├── dist
    │   ├── assets
    │   ├── index.html
    │   └── vite.svg
    ├── eslint.config.js
    ├── index.html
    ├── instructions
    │   └── instructions.md
    ├── package-lock.json
    ├── package.json
    ├── postcss.config.js
    ├── problem.md
    ├── public
    │   └── vite.svg
    ├── server
    │   ├── index.ts
    │   └── tsconfig.json
    ├── src
    │   ├── App.css
    │   ├── App.tsx
    │   ├── assets
    │   ├── components
    │   ├── index.css
    │   ├── lib
    │   ├── main.tsx
    │   ├── services
    │   ├── types
    │   └── vite-env.d.ts
    ├── tailwind.config.js
    ├── tsconfig.app.json
    ├── tsconfig.json
    ├── tsconfig.node.json
    ├── tsconfig.node.tsbuildinfo
    ├── tsconfig.tsbuildinfo
    ├── undefined
    ├── uploads
    ├── vite.config.d.ts
    ├── vite.config.js
    └── vite.config.ts
