import * as ExcelJS from 'exceljs';
import { ExcelConfig, ExcelGroup, ExcelUploadResponse, CellType } from '../types/excel.types';

const DEFAULT_CONFIG: ExcelConfig = {
  sheetName: 'GEOLOGY',
  groupRow: 3,
  columnRow: 4,
  dataStartRow: 5
};

export interface DropdownData {
  columnName: string;
  values: string[];
}

export class ExcelService {
  private workbook: ExcelJS.Workbook | null = null;
  private worksheet: ExcelJS.Worksheet | null = null;
  private currentFile: File | null = null;
  private dropdownData: Map<string, string[]> = new Map();
  private columnTypes: Map<string, CellType> = new Map();
  private validationMap: Map<string, ExcelJS.DataValidation> = new Map();

  /**
   * Excel dosyasını yükler ve GEOLOGY sayfasındaki dropdown hücreleri,
   * DATA sayfasındaki referansları vb. analiz eder.
   */
  async loadExcelFile(file: File): Promise<ExcelUploadResponse> {
    try {
      // 1. Dosya boyutu kontrolü
      if (file.size > 20 * 1024 * 1024) {
        return {
          success: false,
          message: 'Dosya boyutu çok büyük (max: 20MB)',
          error: 'File size too large'
        };
      }

      // 2. Dosyayı chunk'lar halinde oku
      const arrayBuffer = await new Promise<ArrayBuffer>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target?.result as ArrayBuffer);
        reader.onerror = (e) => reject(e);
        reader.readAsArrayBuffer(file);
      });

      console.log('File loaded into memory, processing...');

      // 3. Workbook'u yükle
      this.workbook = new ExcelJS.Workbook();
      await this.workbook.xlsx.load(arrayBuffer);

      // 4. GEOLOGY sayfasını bul
      const worksheet = this.workbook.getWorksheet(DEFAULT_CONFIG.sheetName);
      if (!worksheet) {
        return {
          success: false,
          message: 'GEOLOGY sayfası bulunamadı',
          error: 'Sheet not found'
        };
      }

      console.log('GEOLOGY sheet found, analyzing...');

      this.worksheet = worksheet;
      this.currentFile = file;

      // 5. Performans için önce metadata analizi yap
      const rowCount = worksheet.rowCount;
      const colCount = worksheet.columnCount;

      if (rowCount > 10000 || colCount > 100) {
        return {
          success: false,
          message: 'Dosya çok büyük (max: 10000 satır, 100 sütun)',
          error: 'File too large'
        };
      }

      console.log(`Processing ${rowCount} rows and ${colCount} columns...`);

      // 6. Dropdown hücrelerini tespit et
      await this.detectDropDownCells();
      
      // 7. DATA sayfasından dropdown verilerini al
      await this.extractDropdownData();
      
      // 8. Grupları ve satırları çözümle
      const groups = await this.extractGroups();
      const rows = await this.extractRows(groups);

      console.log('File processing completed successfully');

      return {
        success: true,
        message: 'Excel dosyası başarıyla yüklendi',
        data: {
          groups,
          rows,
          dropdownFields: Array.from(this.dropdownData.keys())
        }
      };
    } catch (error) {
      console.error('Excel yükleme hatası:', error);
      return {
        success: false,
        message: 'Excel dosyası yüklenirken hata oluştu',
        error: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  }

  /**
   * GEOLOGY sayfasındaki hücrelerin dataValidation özelliğini
   * tarayarak type === 'list' olanları (dropdown) tespit eder,
   * formülü parse edip this.dropdownData'ya ekler.
   * Her kolon için columnTypes map'ine 'select' veya 'text' atar.
   */
  private async detectDropDownCells(): Promise<void> {
    if (!this.worksheet) return;

    console.log('Detecting dropdown cells...');
    
    // Sütun başlıklarını al (örneğin 4. satırda kolon başlıkları)
    const headerRow = this.worksheet.getRow(DEFAULT_CONFIG.columnRow);
    const columnNames = new Map<number, string>();
    
    // Sütun isimlerini topla
    headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const columnName = cell.value?.toString();
      if (columnName) {
        columnNames.set(colNumber, columnName);
      }
    });

    // Data satırlarını dön
    this.worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber >= DEFAULT_CONFIG.dataStartRow) {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const columnName = columnNames.get(colNumber);
          if (!columnName) return;

          // Hücrenin dataValidation ayarını al
          const validation = cell.dataValidation;
          console.log(`Checking cell [Row: ${rowNumber}, Col: ${colNumber}] ${columnName}:`, validation);

          // YENİ: Data validation yok veya tip 'list' değilse "text"
          if (!validation) {
            // Data validation yok → text
            if (!this.columnTypes.has(columnName)) {
              this.columnTypes.set(columnName, 'text');
            }
          } else {
            if (validation.type === 'list') {
              // LIST tipinde data validation varsa → dropdown demektir
              console.log(`Found dropdown at [${rowNumber}, ${colNumber}] ${columnName}`);
              console.log('Validation:', validation);
              
              // Bu kolonu dropdown olarak işaretle
              this.columnTypes.set(columnName, 'select');
              this.validationMap.set(columnName, validation);

              // formulae varsa verileri çıkarma girişiminde bulun
              if (validation.formulae && validation.formulae.length > 0) {
                const formula = validation.formulae[0];
                console.log(`Formula for ${columnName}:`, formula);

                if (typeof formula === 'string') {
                  // Data sayfasına referans (ör. =DATA!A2:A10)
                  if (formula.includes('DATA!')) {
                    const range = formula.replace(/^=/, '');
                    this.extractValidationValues(columnName, range);
                  } 
                  // Virgülle ayrılmış doğrudan değerler (ör. =Elma,Armut,Muz)
                  else if (formula.includes(',')) {
                    const values = formula
                      .replace(/^=/, '')
                      .split(',')
                      .map(v => v.trim());
                    this.dropdownData.set(columnName, values);
                  }
                }
              }
            } else {
              // YENİ: Data validation var ama type 'list' değil → text
              if (!this.columnTypes.has(columnName)) {
                this.columnTypes.set(columnName, 'text');
              }
            }
          }
        });
      }
    });

    console.log('Detected column types:', Object.fromEntries(this.columnTypes));
    console.log('Validation map:', Object.fromEntries(this.validationMap));
  }

  /**
   * DATA sayfası mevcutsa oradaki referansları bulup dropdownData map'ine ekler.
   */
  private async extractDropdownData(): Promise<void> {
    if (!this.workbook) return;

    const dataSheet = this.workbook.getWorksheet('DATA');
    if (!dataSheet) {
      console.warn('DATA sayfası bulunamadı');
      return;
    }

    console.log('Processing DATA sheet...');

    // validationMap içinde formulae[0] ile referans (DATA!A2:A10 vs.) olanları işliyoruz
    this.validationMap.forEach((validation, columnName) => {
      if (validation.formulae && validation.formulae[0]) {
        const formula = validation.formulae[0].toString();
        if (formula.includes('DATA!')) {
          const range = formula.replace(/^=/, '');
          this.extractValidationValues(columnName, range);
        }
      }
    });
  }

  /**
   * Belirtilen hücre aralığındaki değerleri alıp this.dropdownData'ya ekler.
   * Örnek range formatı: "DATA!A2:A100"
   */
  private extractValidationValues(columnName: string, range: string): void {
    try {
      if (!this.workbook) return;

      console.log(`Extracting values for ${columnName} from range: ${range}`);

      // Örneğin range = "DATA!A2:A100" 
      const [sheetName, cellRange] = range.split('!');
      const sheet = this.workbook.getWorksheet(sheetName.replace(/^=/, ''));
      
      if (!sheet) {
        console.warn(`Sheet ${sheetName} not found`);
        return;
      }

      const values = new Set<string>();
      
      // A2:A100'ü ayır
      const [startCell, endCell] = cellRange.split(':');
      const startCol = startCell.replace(/[0-9]/g, '');
      const startRow = parseInt(startCell.replace(/[^0-9]/g, ''));
      const endCol = endCell.replace(/[0-9]/g, '');
      const endRow = parseInt(endCell.replace(/[^0-9]/g, ''));

      console.log(`Reading range: ${startCol}${startRow} to ${endCol}${endRow}`);

      for (let row = startRow; row <= endRow; row++) {
        const cell = sheet.getCell(`${startCol}${row}`);
        const value = cell.value;
        if (value !== null && value !== undefined && value !== '') {
          values.add(value.toString().trim());
        }
      }

      if (values.size > 0) {
        const sortedValues = Array.from(values).sort();
        this.dropdownData.set(columnName, sortedValues);
        console.log(`Added ${values.size} values for ${columnName}:`, sortedValues);
      } else {
        console.log(`No values found for ${columnName} in range ${range}`);
      }
    } catch (error) {
      console.error(`Error extracting values for ${columnName}:`, error);
    }
  }

  /**
   * Kolon tipini döndürür (örneğin 'select' ya da 'text').
   */
  public getColumnType(columnName: string): CellType {
    return this.columnTypes.get(columnName) || 'text';
  }

  /**
   * Dropdown için değer listesini döndürür.
   */
  public getDropdownValues(columnName: string): string[] {
    return this.dropdownData.get(columnName) || [];
  }

  /**
   * groupRow satırındaki hücrelerden grup isimlerini, columnRow'daki hücrelerden ise kolon isimlerini okur.
   * Renk bilgilerini de ekleyerek ExcelGroup dizisi döndürür.
   */
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
      
      // Yeni bir grup başlığı bulduk
      if (groupCell.value) {
        if (currentGroup) {
          currentGroup.endColumn = columnIndex - 1;
          groups.push(currentGroup);
        }
        
        // Hücrenin arka plan rengini alma
        let color = 'FFFFFF'; // Varsayılan beyaz
        if (groupCell.fill && 'fgColor' in groupCell.fill && groupCell.fill.fgColor) {
          color = groupCell.fill.fgColor.argb?.substring(2) || 'FFFFFF';
        }
        
        currentGroup = {
          name: groupCell.value.toString(),
          color: color,
          columns: [],
          startColumn: columnIndex,
          endColumn: columnIndex
        };
      }

      // Grup varsa, altına kolon ekle
      if (currentGroup && columnCell.value) {
        currentGroup.columns.push(columnCell.value.toString());
      }

      columnIndex++;
    }

    // Son grup kapat
    if (currentGroup) {
      currentGroup.endColumn = columnIndex - 1;
      groups.push(currentGroup);
    }

    return groups;
  }

  /**
   * dataStartRow'dan itibaren satırları okuyarak, her grup altındaki kolon isimlerine göre değerleri "rowData" şeklinde toplar.
   */
  private async extractRows(groups: ExcelGroup[]): Promise<Record<string, any>[]> {
    if (!this.worksheet) throw new Error('Worksheet not loaded');

    const rows: Record<string, any>[] = [];
    let rowIndex = DEFAULT_CONFIG.dataStartRow;

    while (rowIndex <= this.worksheet.rowCount) {
      const excelRow = this.worksheet.getRow(rowIndex);
      const rowData: Record<string, any> = {};

      groups.forEach(group => {
        for (let i = group.startColumn; i <= group.endColumn; i++) {
          const columnName = this.worksheet!
            .getRow(DEFAULT_CONFIG.columnRow)
            .getCell(i).value?.toString() || '';
          if (columnName) {
            rowData[columnName] = excelRow.getCell(i).value;
          }
        }
      });

      // Boş satır değilse eklensin
      if (Object.keys(rowData).length > 0) {
        rows.push(rowData);
      }

      rowIndex++;
    }

    return rows;
  }

  /**
   * Verilen satır index'ine göre Excel'e veri yazar, workbook'u tekrar kaydeder.
   */
  async saveChanges(rowIndex: number, data: Record<string, any>): Promise<boolean> {
    try {
      if (!this.worksheet || !this.workbook || !this.currentFile) {
        throw new Error('Worksheet not loaded');
      }

      // Veri satırını bul (ör: 5. satır dataStart ise, rowIndex + dataStartRow)
      const row = this.worksheet.getRow(rowIndex + DEFAULT_CONFIG.dataStartRow);

      // Gönderilen data'daki key-value çiftlerini ilgili hücrelere yaz
      Object.entries(data).forEach(([columnName, value]) => {
        const columnIndex = this.findColumnIndex(columnName);
        if (columnIndex !== -1) {
          row.getCell(columnIndex).value = value;
        }
      });

      // Workbook'u buffer'a yaz
      const buffer = await this.workbook.xlsx.writeBuffer();
      // Yeni bir File objesi oluştur
      const newFile = new File([buffer], this.currentFile.name, { type: this.currentFile.type });
      
      // İndirilebilir link oluşturup kullanıcıya sun
      const url = window.URL.createObjectURL(newFile);
      const a = document.createElement('a');
      a.href = url;
      a.download = this.currentFile.name;
      a.click();
      window.URL.revokeObjectURL(url);

      return true;
    } catch (error) {
      console.error('Değişiklikler kaydedilirken hata oluştu:', error);
      return false;
    }
  }

  /**
   * Belirli bir sütun adının Excel'de hangi index'e denk geldiğini bulur.
   */
  private findColumnIndex(columnName: string): number {
    if (!this.worksheet) return -1;

    const columnRow = this.worksheet.getRow(DEFAULT_CONFIG.columnRow);
    let columnIndex = -1;

    for (let i = 1; i <= columnRow.cellCount; i++) {
      if (columnRow.getCell(i).value?.toString() === columnName) {
        columnIndex = i;
        break;
      }
    }

    return columnIndex;
  }
}
