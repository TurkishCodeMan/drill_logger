import React, { useCallback, useState, useMemo } from 'react';
import { useDropzone } from 'react-dropzone';
import { ExcelService } from '../services/excel.service';
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { ExcelData, ExcelGroup } from '@/types/excel.types';
import { FiChevronLeft, FiChevronRight, FiSave } from 'react-icons/fi';
import { toast } from 'sonner';

const excelService = new ExcelService();

export const ExcelForm: React.FC = () => {
  const [excelData, setExcelData] = useState<ExcelData | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [currentRowIndex, setCurrentRowIndex] = useState<number>(0);
  const [formData, setFormData] = useState<Record<string, any>>({});
  const [isSaving, setIsSaving] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [expandedGroups, setExpandedGroups] = useState<Record<string, boolean>>({});

  const onDrop = useCallback(async (acceptedFiles: File[]) => {
    try {
      const file = acceptedFiles[0];
      if (!file) return;

      // Dosya boyutu kontrolü (20MB)
      if (file.size > 20 * 1024 * 1024) {
        setError('Dosya boyutu 20MB\'dan küçük olmalıdır');
        return;
      }

      setIsLoading(true);
      setError(null);

      // Yükleme işlemi başladı bildirimi
      toast.info('Excel dosyası yükleniyor...');

      const result = await excelService.loadExcelFile(file);
      
      if (result.success && result.data) {
        setExcelData(result.data);
        setError(null);
        // İlk satırın verilerini form'a yükle
        if (result.data.rows.length > 0) {
          setFormData(result.data.rows[0]);
          // İlk grubu otomatik aç
          if (result.data.groups.length > 0) {
            setExpandedGroups({ [result.data.groups[0].name]: true });
          }
        }
        toast.success('Excel dosyası başarıyla yüklendi');
      } else {
        setError(result.error || 'Bilinmeyen bir hata oluştu');
        setExcelData(null);
        toast.error(result.error || 'Dosya yüklenirken hata oluştu');
      }
    } catch (error) {
      console.error('Dosya yükleme hatası:', error);
      setError('Dosya yüklenirken hata oluştu');
      toast.error('Dosya yüklenirken beklenmeyen bir hata oluştu');
      setExcelData(null);
    } finally {
      setIsLoading(false);
    }
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    maxSize: 20 * 1024 * 1024, // 20MB
    multiple: false
  });

  const handlePreviousRow = () => {
    if (currentRowIndex > 0 && excelData) {
      setCurrentRowIndex(prev => prev - 1);
      setFormData(excelData.rows[currentRowIndex - 1]);
    }
  };

  const handleNextRow = () => {
    if (excelData && currentRowIndex < excelData.rows.length - 1) {
      setCurrentRowIndex(prev => prev + 1);
      setFormData(excelData.rows[currentRowIndex + 1]);
    }
  };

  const formatValue = (value: any): string => {
    if (value === null || value === undefined) return '';
    
    if (typeof value === 'object') {
      // Excel formül objesi kontrolü
      if ('formula' in value && 'result' in value) {
        return value.result || '';
      }
      // Sadece sonuç değeri varsa
      if ('result' in value) {
        return value.result || '';
      }
      // Sadece text değeri varsa
      if ('text' in value) {
        return value.text || '';
      }
      // Sadece value değeri varsa
      if ('value' in value) {
        return value.value || '';
      }
      
      // Boş obje kontrolü
      if (Object.keys(value).length === 0) {
        return '';
      }
      
      // Diğer durumlar için ilk anlamlı değeri bul
      for (const key of ['result', 'text', 'value', 'formula']) {
        if (key in value && value[key]) {
          return String(value[key]);
        }
      }
      
      // Hiçbir anlamlı değer bulunamazsa boş string döndür
      return '';
    }
    
    // Obje değilse string'e çevir
    return String(value || '');
  };

  const handleInputChange = (columnName: string, value: string) => {
    setFormData(prev => ({
      ...prev,
      [columnName]: value
    }));
  };

  const handleSave = async () => {
    if (!excelData) return;

    setIsSaving(true);
    try {
      const success = await excelService.saveChanges(currentRowIndex, formData);
      
      if (success) {
        const updatedRows = [...excelData.rows];
        updatedRows[currentRowIndex] = formData;
        setExcelData({
          ...excelData,
          rows: updatedRows
        });
        
        toast.success('Değişiklikler kaydedildi');
        
        if (currentRowIndex < excelData.rows.length - 1) {
          handleNextRow();
        }
      } else {
        toast.error('Değişiklikler kaydedilemedi');
      }
    } catch (error) {
      toast.error('Bir hata oluştu');
    } finally {
      setIsSaving(false);
    }
  };

  const toggleGroup = (groupName: string) => {
    setExpandedGroups(prev => ({
      ...prev,
      [groupName]: !prev[groupName]
    }));
  };

  // Grupları birleştir
  const mergedGroups = useMemo(() => {
    if (!excelData?.groups) return [];

    // İzin verilen MINERALIZATION alanları - tam olarak belirtilen formatta
    const allowedMineralizationColumns = [
      'Vein          Type',
      'QzVn %',
      'QzBx %',
      'Bx                   Clast',
      'Bx                  Matrix',
      'Min                    Zone',
      'Pima                   Sample',
      'Comments'
    ];

    // İzin verilen ALTERATION alanları - tam olarak belirtilen formatta
    const allowedAlterationColumns = [
      'Alteration               Type',
      'Sil            Deg',
      'Pyrite',
      'Vuggy%',
      'Arg                       Kaolinite',
      'Serisite',
      'Dickite',
      'Alunite',
      'Chlorite',
      'Epidote',
      'Carbonate',
      'Gypsum',
      'Oxidation'
    ];

    // İzin verilen GEOTECHNICAL alanları - tam olarak belirtilen formatta
    const allowedGeotechnicalColumns = [
      'Weathering',
      'recovery_m',
      'rqd',
      'Mjr_defect              type',
      'Mjr_defect               alpha',
      'Frac',
      'Rock                      Strength'
    ];

    const groupMap = new Map<string, ExcelGroup>();

    excelData.groups.forEach(group => {
      // Grup adını kontrol et ve düzelt
      const groupName = group.name && group.name !== '[object Object]' ? group.name : 'INFO';
      
      if (groupMap.has(groupName)) {
        // Mevcut grubun sütunlarını ekle
        const existingGroup = groupMap.get(groupName)!;
        const uniqueColumns = new Set([...existingGroup.columns, ...group.columns]);
        
        let filteredColumns = Array.from(uniqueColumns);
        
        // MINERALIZATION grubu için tam eşleşme filtrelemesi
        if (groupName.toLowerCase().includes('mineralization')) {
          filteredColumns = filteredColumns.filter(column => 
            allowedMineralizationColumns.includes(column)
          );
        }
        // ALTERATION grubu için tam eşleşme filtrelemesi
        else if (groupName.toLowerCase().includes('alteration')) {
          filteredColumns = filteredColumns.filter(column => 
            allowedAlterationColumns.includes(column)
          );
        }
        // GEOTECHNICAL grubu için tam eşleşme filtrelemesi
        else if (groupName.toLowerCase().includes('geotecnical')) {
          filteredColumns = filteredColumns.filter(column => 
            allowedGeotechnicalColumns.includes(column)
          );
        }
        
        existingGroup.columns = filteredColumns;
      } else {
        // Yeni grup oluştur
        let columns = [...group.columns];
        
        // MINERALIZATION grubu için tam eşleşme filtrelemesi
        if (groupName.toLowerCase().includes('mineralization')) {
          columns = columns.filter(column => 
            allowedMineralizationColumns.includes(column)
          );
        }
        // ALTERATION grubu için tam eşleşme filtrelemesi
        else if (groupName.toLowerCase().includes('alteration')) {
          columns = columns.filter(column => 
            allowedAlterationColumns.includes(column)
          );
        }
        // GEOTECHNICAL grubu için tam eşleşme filtrelemesi
        else if (groupName.toLowerCase().includes('geotechnical')) {
          columns = columns.filter(column => 
            allowedGeotechnicalColumns.includes(column)
          );
        }
        
        groupMap.set(groupName, { ...group, name: groupName, columns });
      }
    });

    return Array.from(groupMap.values());
  }, [excelData?.groups]);

  // LithCode seçenekleri
  const lithCodes = [
    'ALL',
    'ABX',
    'AND',
    'CORELOST',
    'FZ',
    'FZ1',
    'HBX',
    'NC',
    'MNBX',
    'MS',
    'PLBX',
    'PQRF1',
    'VOLSED',
    'VBX',
    'QBX',
    'QV',
    'QZSULP'
  ];

  const renderFormField = (column: string) => {
    const columnType = excelService.getColumnType(column);
    const dropdownValues = excelService.getDropdownValues(column);

    // LithCode sütunları için özel dropdown
    if (column === 'Lith1               Code' || column === 'Lith2               Code') {
      const currentValue = formatValue(formData[column]);
      
      return (
        <div className="relative">
          <Select
            defaultValue={currentValue || "ALL"}
            value={currentValue || "ALL"}
            onValueChange={(value) => handleInputChange(column, value)}
            disabled={isSaving}
          >
            <SelectTrigger className="h-6 text-[10px]">
              <SelectValue placeholder="LithCode seçiniz..." />
            </SelectTrigger>
            <SelectContent>
              {lithCodes.map((code) => (
                <SelectItem key={code} value={code} className="text-[10px]">
                  {code}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
          <div className="absolute top-0 right-0 -mr-3 text-[8px] text-muted-foreground">
            (L)
          </div>
        </div>
      );
    }

    if (columnType === 'select' && dropdownValues.length > 0) {
      const currentValue = formatValue(formData[column]);
      
      return (
        <div className="relative">
          <Select
            defaultValue={currentValue}
            value={currentValue}
            onValueChange={(value) => handleInputChange(column, value)}
            disabled={isSaving}
          >
            <SelectTrigger className="h-6 text-[10px]">
              <SelectValue placeholder="Seçiniz..." />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="default">Seçiniz...</SelectItem>
              {dropdownValues.map((value, index) => (
                <SelectItem key={`${value}-${index}`} value={value} className="text-[10px]">
                  {value}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
          <div className="absolute top-0 right-0 -mr-3 text-[8px] text-muted-foreground">
            (L)
          </div>
        </div>
      );
    }

    return (
      <Input
        id={column}
        value={formatValue(formData[column])}
        onChange={(e) => handleInputChange(column, e.target.value)}
        disabled={isSaving}
        className="h-6 text-[10px]"
      />
    );
  };

  return (
    <div className="h-screen">
      {!excelData ? (
        <Card className="p-6">
          <div className="space-y-4">
            <div>
              <h2 className="text-2xl font-bold mb-4">Excel Dosyası Yükle</h2>
              <p className="text-gray-600 mb-4">
                Lütfen GEOLOGY sayfası içeren bir Excel dosyası yükleyin.
                <br />
                <span className="text-sm text-muted-foreground">
                  (Maksimum dosya boyutu: 20MB)
                </span>
              </p>
            </div>

            <div
              {...getRootProps()}
              className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors relative
                ${isDragActive ? 'border-primary bg-primary/10' : 'border-gray-300 hover:border-primary'}`}
            >
              <input {...getInputProps()} />
              <div className="space-y-2">
                <div className="text-4xl mb-4">📊</div>
                {isLoading ? (
                  <div className="space-y-2">
                    <div className="animate-pulse">Dosya yükleniyor...</div>
                    <div className="text-sm text-muted-foreground">
                      Lütfen bekleyin, bu işlem biraz zaman alabilir
                    </div>
                  </div>
                ) : isDragActive ? (
                  <p className="text-primary">Dosyayı buraya bırakın...</p>
                ) : (
                  <>
                    <p className="text-gray-600">
                      Dosyayı sürükleyip bırakın veya seçmek için tıklayın
                    </p>
                    <p className="text-sm text-gray-500">
                      (Sadece .xlsx ve .xls dosyaları desteklenir)
                    </p>
                  </>
                )}
              </div>
            </div>

            {error && (
              <div className="bg-destructive/10 text-destructive px-4 py-2 rounded-md text-sm">
                {error}
              </div>
            )}
          </div>
        </Card>
      ) : (
        <div className="h-screen flex flex-col">
          <div className="flex-1 relative overflow-hidden">
            <div className="sticky top-0 z-10 bg-background border-b py-1 px-2">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-1">
                  <span className="text-sm font-bold">Satır {currentRowIndex + 1} / {excelData.rows.length}</span>
                </div>
                <Button 
                  onClick={handleSave} 
                  size="sm"
                  disabled={isSaving}
                  className="h-6 px-2 flex items-center gap-1 text-xs"
                >
                  <FiSave size={12} />
                  {isSaving ? 'Kaydediliyor...' : 'Kaydet'}
                </Button>
              </div>
            </div>
            <div className="overflow-auto h-[calc(100vh-3rem)]">
              <div className="space-y-0.5">
                {mergedGroups.map((group, groupIndex) => (
                  <div key={groupIndex} className="bg-card">
                    <div className="bg-muted/50 px-2 py-1 border-b flex items-center gap-1">
                      <div 
                        className="w-2 h-2 rounded-full flex-shrink-0" 
                        style={{ backgroundColor: `#${group.color}` }}
                      />
                      <h3 className="text-xs font-medium text-muted-foreground">
                        {group.name}
                        <span className="ml-1 text-[10px]">
                          ({group.columns.length})
                        </span>
                      </h3>
                    </div>
                    <div className="p-0.5">
                      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-0.5">
                        {group.columns.sort().map((column, columnIndex) => (
                          <div key={columnIndex} className="min-w-[120px]">
                            <Label htmlFor={column} className="text-[10px] text-muted-foreground mb-0.5 block truncate" title={column}>
                              {column}
                            </Label>
                            {renderFormField(column)}
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
          
          {/* Sayfalama Butonları - Sağ Alt Sabit */}
          <div className="fixed bottom-1 right-1 flex items-center gap-1 bg-background/80 backdrop-blur-sm p-1 rounded-md shadow-lg z-50">
            <Button
              variant="outline"
              size="sm"
              onClick={handlePreviousRow}
              disabled={currentRowIndex === 0 || isSaving}
              className="h-6 w-6 p-0"
            >
              <FiChevronLeft size={12} />
            </Button>
            <span className="text-xs font-medium px-1">
              {currentRowIndex + 1} / {excelData.rows.length}
            </span>
            <Button
              variant="outline"
              size="sm"
              onClick={handleNextRow}
              disabled={currentRowIndex === excelData.rows.length - 1 || isSaving}
              className="h-6 w-6 p-0"
            >
              <FiChevronRight size={12} />
            </Button>
          </div>
        </div>
      )}
    </div>
  );
}; 