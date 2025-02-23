import * as ExcelJS from 'exceljs';
import { ExcelConfig, ExcelGroup, ExcelUploadResponse, CellType } from '../types/excel.types';

const DEFAULT_CONFIG: ExcelConfig = {
  sheetName: 'GEOLOGY',
  groupRow: 3,
  columnRow: 4,
  dataStartRow: 5
};

const COLLAR_CONFIG: ExcelConfig = {
  sheetName: 'COLLAR',
  groupRow: 1,
  columnRow: 1,
  dataStartRow: 2
};

export interface DropdownData {
  columnName: string;
  values: string[];
}

export class ExcelService {
  private workbook: ExcelJS.Workbook | null = null;
  private worksheet: ExcelJS.Worksheet | null = null;
  private collarWorksheet: ExcelJS.Worksheet | null = null;
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

      // 4. GEOLOGY ve COLLAR sayfalarını bul
      const worksheet = this.workbook.getWorksheet(DEFAULT_CONFIG.sheetName);
      const collarWorksheet = this.workbook.getWorksheet(COLLAR_CONFIG.sheetName);
      
      if (!worksheet) {
        return {
          success: false,
          message: 'GEOLOGY sayfası bulunamadı',
          error: 'Sheet not found'
        };
      }

      console.log('GEOLOGY sheet found, analyzing...');
      console.log('COLLAR sheet found:', !!collarWorksheet);

      this.worksheet = worksheet;
      this.collarWorksheet = collarWorksheet as any;
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
      
      // 9. COLLAR verilerini çözümle
      const collarData = collarWorksheet ? await this.extractCollarData() : null;

      console.log('File processing completed successfully');

      return {
        success: true,
        message: 'Excel dosyası başarıyla yüklendi',
        data: {
          groups,
          rows,
          dropdownFields: Array.from(this.dropdownData.keys()),
          collarData
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
    
    // Sütun başlıklarını al
    const headerRow = this.worksheet.getRow(DEFAULT_CONFIG.columnRow);
    const columnNames = new Map<number, string>();
    
    // Sütun isimlerini topla
    headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const columnName = cell.value?.toString();
      if (columnName) {
        columnNames.set(colNumber, columnName);
      }
    });

    // Zone, Weathering, Mjr_defect type, Rock Strength, Alteration Type, Arg Kaolinite, diğer ALTERATION alanları ve Min Zone için özel işlem
    columnNames.forEach((columnName, __) => {
      if (columnName.includes('Zone') || 
          columnName === 'Weathering' || 
          columnName === 'Mjr_defect              type' ||
          columnName === 'Rock                      Strength' ||
          columnName === 'Alteration               Type' ||
          columnName === 'Arg                       Kaolinite' ||
          columnName === 'Serisite' ||
          columnName === 'Dickite' ||
          columnName === 'Alunite' ||
          columnName === 'Chlorite' ||
          columnName === 'Epidote' ||
          columnName === 'Carbonate' ||
          columnName === 'Oxidation' ||
          columnName === 'Min                    Zone' ||
          columnName === 'sample_this') {
        this.columnTypes.set(columnName, 'select');
        
        // sample_this için özel dropdown değerleri
        if (columnName === 'sample_this') {
          this.dropdownData.set(columnName, ['Yes', 'No']);
        }
      }
    });

    // Data satırlarını dön
    this.worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber >= DEFAULT_CONFIG.dataStartRow) {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const columnName = columnNames.get(colNumber);
          if (!columnName) return;

          // Zone, Weathering, Mjr_defect type, Arg Kaolinite, diğer ALTERATION alanları ve Min Zone için özel kontrol ekle
          if (columnName.includes('Zone') || 
              columnName === 'Weathering' || 
              columnName === 'Mjr_defect              type' ||
              columnName === 'Arg                       Kaolinite' ||
              columnName === 'Serisite' ||
              columnName === 'Dickite' ||
              columnName === 'Alunite' ||
              columnName === 'Chlorite' ||
              columnName === 'Epidote' ||
              columnName === 'Carbonate' ||
              columnName === 'Oxidation' ||
              columnName === 'Min                    Zone') {
            return;
          }

          // Diğer hücreler için normal kontrol
          const validation = cell.dataValidation;
          if (!validation) {
            if (!this.columnTypes.has(columnName)) {
              this.columnTypes.set(columnName, 'text');
            }
          } else {
            if (validation.type === 'list') {
              this.columnTypes.set(columnName, 'select');
              this.validationMap.set(columnName, validation);

              if (validation.formulae && validation.formulae.length > 0) {
                const formula = validation.formulae[0];
                if (typeof formula === 'string') {
                  if (formula.includes('DATA!')) {
                    const range = formula.replace(/^=/, '');
                    this.extractValidationValues(columnName, range);
                  } else if (formula.includes(',')) {
                    const values = formula
                      .replace(/^=/, '')
                      .split(',')
                      .map(v => v.trim());
                    this.dropdownData.set(columnName, values);
                  }
                }
              }
            } else {
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

    // INFO grubu
    const infoGroup: ExcelGroup = {
      name: 'INFO',
      color: 'FFFFFF',
      columns: ['E1_HEADER', 'E2_INPUT'],
      startColumn: 1,
      endColumn: 2
    };
    groups.push(infoGroup);

    // SAMPLE grubu
    const sampleGroup: ExcelGroup = {
      name: 'SAMPLE',
      color: 'FFFFFF',
      columns: ['sample_this', 'Sample Number'],
      startColumn: 3,
      endColumn: 4
    };
    groups.push(sampleGroup);

    // CG4 ve CH4 başlıklarını ve verileri işle
    //const cg4Value = this.worksheet.getCell('CG4').value?.toString() || 'sample_this';
    //const ch4Value = this.worksheet.getCell('CH4').value?.toString() || 'Sample Number';

    // CG5 ve CH5'ten başlayan verileri al
    let rowIndex = 5;
    while (rowIndex <= this.worksheet.rowCount) {
      const cgValue = this.worksheet.getCell(`CG${rowIndex}`).value;
      const chValue = this.worksheet.getCell(`CH${rowIndex}`).value;
      if (cgValue || chValue) {
        // Verileri formData'ya ekle
         // const rowData: Record<string, any> = {
          //'sample_this': cgValue,
          //'Sample Number': chValue
        //};
        // ... existing row processing code ...
      }
      rowIndex++;
    }

    while (columnIndex <= columnRow.cellCount) {
      const groupCell = groupRow.getCell(columnIndex);
      const columnCell = columnRow.getCell(columnIndex);
      
      // Yeni bir grup başlığı bulduk ve INFO değilse
      if (groupCell.value && groupCell.value.toString().toLowerCase() !== 'info') {
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

      // Grup varsa ve INFO değilse, altına kolon ekle
      if (currentGroup && columnCell.value && currentGroup.name.toLowerCase() !== 'info') {
        currentGroup.columns.push(columnCell.value.toString());
      }

      columnIndex++;
    }

    // Son grup kapat
    if (currentGroup && currentGroup.name.toLowerCase() !== 'info') {
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

    // E1 ve E2 değerlerini al
    const e1Value = this.worksheet.getCell('E1').value?.toString() || 'Başlık';
    const e2Value = this.worksheet.getCell('E2').value?.toString() || '';

    while (rowIndex <= this.worksheet.rowCount) {
      const excelRow = this.worksheet.getRow(rowIndex);
      const rowData: Record<string, any> = {
        'E1_HEADER': e1Value,
        'E2_INPUT': e2Value,
        'sample_this': this.worksheet.getCell(`CG${rowIndex}`).value,
        'Sample Number': this.worksheet.getCell(`CH${rowIndex}`).value
      };

      groups.forEach(group => {
        if (group.name.toLowerCase() !== 'info') {
          for (let i = group.startColumn; i <= group.endColumn; i++) {
            const columnName = this.worksheet!
              .getRow(DEFAULT_CONFIG.columnRow)
              .getCell(i).value?.toString() || '';
            if (columnName) {
              rowData[columnName] = excelRow.getCell(i).value;
            }
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
  async saveChanges(rowIndex: number, data: Record<string, any>, collarData?: Record<string, any>): Promise<boolean> {
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

      // COLLAR verilerini güncelle
      if (collarData && this.collarWorksheet) {
//        const collarRow = this.collarWorksheet.getRow(COLLAR_CONFIG.dataStartRow);
        Object.entries(collarData).forEach(([columnName, value]) => {
          const cell = this.findCollarCell(columnName);
          if (cell) {
            cell.value = value;
          }
        });
      }

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

  /**
   * COLLAR sayfasında belirli bir sütun adının hücresini bulur.
   */
  private findCollarCell(columnName: string): ExcelJS.Cell | null {
    if (!this.collarWorksheet) return null;

    const headerRow = this.collarWorksheet.getRow(COLLAR_CONFIG.columnRow);
    let columnIndex = -1;

    headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (cell.value?.toString() === columnName) {
        columnIndex = colNumber;
      }
    });

    if (columnIndex === -1) return null;

    return this.collarWorksheet.getCell(COLLAR_CONFIG.dataStartRow, columnIndex);
  }

  private async extractCollarData(): Promise<any> {
    if (!this.collarWorksheet) return null;

    const headerRow = this.collarWorksheet.getRow(COLLAR_CONFIG.columnRow);
    const dataRow = this.collarWorksheet.getRow(COLLAR_CONFIG.dataStartRow);
    const collarData: Record<string, any> = {};

    headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const columnName = cell.value?.toString();
      if (columnName) {
        collarData[columnName] = dataRow.getCell(colNumber).value;
      }
    });

    return collarData;
  }

  async getLithCodesFromDataSheet(): Promise<string[]> {
    try {
      if (!this.workbook) {
        throw new Error('Excel dosyası yüklenmemiş');
      }

      // Data Sheet'i al
      const dataSheet = this.workbook.getWorksheet('DATA');
      if (!dataSheet) {
        throw new Error('Data Sheet bulunamadı');
      }

      const lithCodes: string[] = [];
      let rowIndex = 2; // A2'den başla
      
      // A sütunundaki değerleri oku
      while (rowIndex <= 20) { // A20'ye kadar
        const cell = dataSheet.getCell(`A${rowIndex}`);
        
        // Eğer hücre koyu renkliyse veya boşsa döngüyü bitir
        if (!cell || cell.font?.bold || !cell.value) {
          break;
        }
        
        // Değeri ekle
        lithCodes.push(cell.value.toString().trim());
        rowIndex++;
      }

      return lithCodes;
    } catch (error) {
      console.error('LithCode değerleri okunurken hata:', error);
      return [];
    }
  }

  async getZoneCodesFromDataSheet(): Promise<string[]> {
    try {
      if (!this.workbook) {
        throw new Error('Excel dosyası yüklenmemiş');
      }

      // Data Sheet'i al
      const dataSheet = this.workbook.getWorksheet('DATA');
      if (!dataSheet) {
        throw new Error('Data Sheet bulunamadı');
      }

      const zoneCodes: string[] = [];
      let rowIndex = 29; // A29'dan başla
      
      // A sütunundaki değerleri oku
      while (rowIndex <= 40) { // A40'a kadar
        const cell = dataSheet.getCell(`A${rowIndex}`);
        
        // Eğer hücre boşsa döngüyü bitir
        if (!cell || !cell.value) {
          break;
        }
        
        // Değeri ekle
        zoneCodes.push(cell.value.toString().trim());
        rowIndex++;
      }

      return zoneCodes;
    } catch (error) {
      console.error('Zone değerleri okunurken hata:', error);
      return [];
    }
  }

  async getWeatheringCodesFromDataSheet(): Promise<string[]> {
    try {
      if (!this.workbook) {
        throw new Error('Excel dosyası yüklenmemiş');
      }

      // Data Sheet'i al
      const dataSheet = this.workbook.getWorksheet('DATA');
      if (!dataSheet) {
        throw new Error('Data Sheet bulunamadı');
      }

      const weatheringCodes: string[] = [];
      let rowIndex = 2; // M2'den başla
      
      // M sütunundaki değerleri oku
      while (rowIndex <= 30) { // M30'a kadar
        const cell = dataSheet.getCell(`M${rowIndex}`);
        
        // Eğer hücre boşsa döngüyü bitir
        if (!cell || !cell.value) {
          break;
        }
        
        // Değeri ekle
        weatheringCodes.push(cell.value.toString().trim());
        rowIndex++;
      }

      return weatheringCodes;
    } catch (error) {
      console.error('Weathering değerleri okunurken hata:', error);
      return [];
    }
  }

  async getMjrDefectTypesFromDataSheet(): Promise<string[]> {
    try {
      if (!this.workbook) {
        throw new Error('Excel dosyası yüklenmemiş');
      }

      // Data Sheet'i al
      const dataSheet = this.workbook.getWorksheet('DATA');
      if (!dataSheet) {
        throw new Error('Data Sheet bulunamadı');
      }

      const defectTypes: string[] = [];
      let rowIndex = 33; // Q33'den başla
      
      // Q sütunundaki değerleri oku
      while (rowIndex <= 50) { // Q50'ye kadar
        const cell = dataSheet.getCell(`Q${rowIndex}`);
        
        // Eğer hücre boşsa döngüyü bitir
        if (!cell || !cell.value) {
          break;
        }
        
        // Değeri ekle
        defectTypes.push(cell.value.toString().trim());
        rowIndex++;
      }

      return defectTypes;
    } catch (error) {
      console.error('Mjr_defect type değerleri okunurken hata:', error);
      return [];
    }
  }

  async getRockStrengthFromDataSheet(): Promise<string[]> {
    try {
      if (!this.workbook) {
        throw new Error('Excel dosyası yüklenmemiş');
      }

      // Data Sheet'i al
      const dataSheet = this.workbook.getWorksheet('DATA');
      if (!dataSheet) {
        throw new Error('Data Sheet bulunamadı');
      }

      const strengthValues: string[] = [];
      let rowIndex = 12; // W12'den başla
      
      // W sütunundaki değerleri oku
      while (rowIndex <= 30) { // W30'a kadar
        const cell = dataSheet.getCell(`W${rowIndex}`);
        
        // Eğer hücre boşsa döngüyü bitir
        if (!cell || !cell.value) {
          break;
        }
        
        // Değeri ekle
        strengthValues.push(cell.value.toString().trim());
        rowIndex++;
      }

      return strengthValues;
    } catch (error) {
      console.error('Rock Strength değerleri okunurken hata:', error);
      return [];
    }
  }

  async getAlterationTypesFromDataSheet(): Promise<string[]> {
    try {
      if (!this.workbook) {
        throw new Error('Excel dosyası yüklenmemiş');
      }

      // Data Sheet'i al
      const dataSheet = this.workbook.getWorksheet('DATA');
      if (!dataSheet) {
        throw new Error('Data Sheet bulunamadı');
      }

      const alterationTypes: string[] = [];
      let rowIndex = 2; // Z2'den başla
      
      // Z sütunundaki değerleri oku
      while (rowIndex <= 40) { // Z40'a kadar
        const cell = dataSheet.getCell(`Z${rowIndex}`);
        
        // Eğer hücre boşsa döngüyü bitir
        if (!cell || !cell.value) {
          break;
        }
        
        // Değeri ekle
        alterationTypes.push(cell.value.toString().trim());
        rowIndex++;
      }

      return alterationTypes;
    } catch (error) {
      console.error('Alteration Type değerleri okunurken hata:', error);
      return [];
    }
  }

  async getMinZoneFromDataSheet(): Promise<string[]> {
    try {
      if (!this.workbook) {
        throw new Error('Excel dosyası yüklenmemiş');
      }

      // Data Sheet'i al
      const dataSheet = this.workbook.getWorksheet('DATA');
      if (!dataSheet) {
        throw new Error('Data Sheet bulunamadı');
      }

      const minZoneValues: string[] = [];
      let rowIndex = 2; // AC2'den başla
      
      // AC sütunundaki değerleri oku
      while (rowIndex <= 30) { // AC30'a kadar
        const cell = dataSheet.getCell(`AC${rowIndex}`);
        
        // Eğer hücre boşsa döngüyü bitir
        if (!cell || !cell.value) {
          break;
        }
        
        // Değeri ekle
        minZoneValues.push(cell.value.toString().trim());
        rowIndex++;
      }

      return minZoneValues;
    } catch (error) {
      console.error('Min Zone değerleri okunurken hata:', error);
      return [];
    }
  }
}
