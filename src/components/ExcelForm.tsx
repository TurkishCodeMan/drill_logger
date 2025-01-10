import React, { useCallback, useState, useMemo } from 'react';
import { useDropzone } from 'react-dropzone';
import { ExcelService } from '../services/excel.service';
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { ExcelData, ExcelGroup } from '@/types/excel.types';
import { ChevronLeft, ChevronRight, Save, ChevronDown, ChevronUp } from 'lucide-react';
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

    const groupMap = new Map<string, ExcelGroup>();

    excelData.groups.forEach(group => {
      if (groupMap.has(group.name)) {
        // Mevcut grubun sütunlarını ekle
        const existingGroup = groupMap.get(group.name)!;
        const uniqueColumns = new Set([...existingGroup.columns, ...group.columns]);
        existingGroup.columns = Array.from(uniqueColumns);
      } else {
        // Yeni grup oluştur
        groupMap.set(group.name, { ...group, columns: [...group.columns] });
      }
    });

    return Array.from(groupMap.values());
  }, [excelData?.groups]);

  const renderFormField = (column: string) => {
    const columnType = excelService.getColumnType(column);
    const dropdownValues = excelService.getDropdownValues(column);
    
    console.log(`Rendering field for ${column}:`, {
      columnType,
      dropdownValues,
      currentValue: formData[column]
    });

    if (columnType === 'select' && dropdownValues.length > 0) {
      console.log(`Rendering SELECT for ${column} with ${dropdownValues.length} options`);
      const currentValue = formatValue(formData[column]);
      
      return (
        <div className="relative">
          <Select
            defaultValue={currentValue}
            value={currentValue}
            onValueChange={(value) => handleInputChange(column, value)}
            disabled={isSaving}
          >
            <SelectTrigger className="h-12 text-lg w-full">
              <SelectValue placeholder={`${column} seçiniz...`} />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="">Seçiniz...</SelectItem>
              {dropdownValues.map((value, index) => (
                <SelectItem key={`${value}-${index}`} value={value}>
                  {value}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
          {columnType === 'select' && (
            <div className="absolute top-0 right-0 -mr-6 text-xs text-muted-foreground">
              (Liste)
            </div>
          )}
        </div>
      );
    }

    console.log(`Rendering INPUT for ${column}`);
    return (
      <Input
        id={column}
        value={formatValue(formData[column])}
        onChange={(e) => handleInputChange(column, e.target.value)}
        disabled={isSaving}
        className="h-12 text-lg"
      />
    );
  };

  const renderForm = () => {
    if (!excelData?.groups.length) return null;

    return (
      <div className="mt-8 pb-20">
        <Card className="relative">
          <CardHeader className="sticky top-0 z-10 bg-card border-b">
            <CardTitle className="flex items-center justify-between">
              <div className="flex items-center gap-4">
                <span className="text-2xl font-bold">Satır {currentRowIndex + 1} / {excelData.rows.length}</span>
                <div className="flex items-center gap-2">
                  <Button
                    variant="outline"
                    size="lg"
                    onClick={handlePreviousRow}
                    disabled={currentRowIndex === 0 || isSaving}
                    className="h-12 w-12"
                  >
                    <ChevronLeft className="h-6 w-6" />
                  </Button>
                  <Button
                    variant="outline"
                    size="lg"
                    onClick={handleNextRow}
                    disabled={currentRowIndex === excelData.rows.length - 1 || isSaving}
                    className="h-12 w-12"
                  >
                    <ChevronRight className="h-6 w-6" />
                  </Button>
                </div>
              </div>
              <Button 
                onClick={handleSave} 
                size="lg"
                disabled={isSaving}
                className="h-12 px-6"
              >
                <Save className="w-6 h-6 mr-2" />
                {isSaving ? 'Kaydediliyor...' : 'Kaydet'}
              </Button>
            </CardTitle>
          </CardHeader>
          <CardContent className="p-6">
            <div className="space-y-6">
              {mergedGroups.map((group, groupIndex) => (
                <Card key={groupIndex} className="overflow-hidden">
                  <button
                    onClick={() => toggleGroup(group.name)}
                    className={`w-full p-4 flex items-center justify-between text-left transition-colors
                      ${expandedGroups[group.name] ? 'bg-primary/5' : 'hover:bg-primary/5'}`}
                  >
                    <div className="flex items-center gap-3">
                      <div 
                        className="w-6 h-6 rounded-full" 
                        style={{ backgroundColor: `#${group.color}` }}
                      />
                      <h3 className="text-xl font-semibold">{group.name}</h3>
                      <span className="text-sm text-muted-foreground">
                        ({group.columns.length} alan)
                      </span>
                    </div>
                    {expandedGroups[group.name] ? (
                      <ChevronUp className="h-6 w-6" />
                    ) : (
                      <ChevronDown className="h-6 w-6" />
                    )}
                  </button>
                  {expandedGroups[group.name] && (
                    <div className="p-4 border-t">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        {group.columns.sort().map((column, columnIndex) => (
                          <div key={columnIndex} className="space-y-3">
                            <Label htmlFor={column} className="text-base">
                              {column}
                            </Label>
                            {renderFormField(column)}
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </Card>
              ))}
            </div>
          </CardContent>
        </Card>
      </div>
    );
  };

  return (
    <div className="container mx-auto p-4">
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
        renderForm()
      )}
    </div>
  );
}; 