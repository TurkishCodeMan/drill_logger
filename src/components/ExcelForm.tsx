import React, { useCallback, useState, useMemo, useEffect } from 'react';
import { useDropzone } from 'react-dropzone';
import { ExcelService } from '../services/excel.service';
import { Button } from "@/components/ui/button"
import { Card } from "@/components/ui/card"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { ExcelData, ExcelGroup } from '@/types/excel.types';
import { FiChevronLeft, FiChevronRight, FiSave, FiMove } from 'react-icons/fi';
import { toast } from 'sonner';
import { DndProvider, useDrag, useDrop } from 'react-dnd';
import { HTML5Backend } from 'react-dnd-html5-backend';

const excelService = new ExcelService();

interface DraggableGroupBoxProps {
  group: ExcelGroup;
  index: number;
  moveBox: (dragIndex: number, hoverIndex: number) => void;
  renderFormField: (column: string) => React.ReactNode;
  style: React.CSSProperties;
}

interface DragItem {
  id: string;
  index: number;
  type: string;
}

// Taşınabilir grup bileşeni
const DraggableGroupBox: React.FC<DraggableGroupBoxProps> = ({ 
  group, 
  index, 
  moveBox, 
  renderFormField, 
  style
}) => {
  const ref = React.useRef<HTMLDivElement>(null);
  const isInfo = group.name.toLowerCase() === 'info';
  const isLithology = group.name.toLowerCase().includes('lithology');
  const isMineralization = group.name.toLowerCase().includes('mineralization');
  const isAlteration = group.name.toLowerCase().includes('alteration');
  const isGeotechnical = group.name.toLowerCase().includes('geotecnical');

  const [{ isDragging }, drag] = useDrag(() => ({
    type: 'group-box',
    item: { 
      id: group.name, 
      index,
      type: 'group-box'
    } as DragItem,
    collect: (monitor) => ({
      isDragging: monitor.isDragging(),
    }),
  }));

  const [, drop] = useDrop(() => ({
    accept: 'group-box',
    hover: (item: DragItem, monitor) => {
      if (!ref.current) return;
      
      const dragIndex = item.index;
      const hoverIndex = index;

      if (dragIndex === hoverIndex) return;

      const hoverBoundingRect = ref.current.getBoundingClientRect();
      const hoverMiddleY = (hoverBoundingRect.bottom - hoverBoundingRect.top) / 2;
      const clientOffset = monitor.getClientOffset();
      const hoverClientY = clientOffset!.y - hoverBoundingRect.top;

      if (dragIndex < hoverIndex && hoverClientY < hoverMiddleY) return;
      if (dragIndex > hoverIndex && hoverClientY > hoverMiddleY) return;

      moveBox(dragIndex, hoverIndex);
      item.index = hoverIndex;
    },
  }));

  drag(drop(ref));

  const getHeaderColor = () => {
    if (isLithology) return '#e6f3ff';
    if (isAlteration) return '#fff2e6';
    if (isGeotechnical) return '#dcfce7';
    if (isMineralization) return '#e6f3ff';
    if (isInfo) return '#f1f5f9';
    if (group.name.toLowerCase() === 'sample') return '#e0f2fe';
    return '#f3e6ff';
  };

  const getBackgroundColor = () => {
    if (isLithology) return 'rgba(230, 243, 255, 0.3)';
    if (isAlteration) return 'rgba(255, 237, 213, 0.3)';
    if (isGeotechnical) return 'rgba(220, 252, 231, 0.3)';
    if (isMineralization) return 'rgba(230, 243, 255, 0.3)';
    if (isInfo) return 'rgba(241, 245, 249, 0.3)';
    if (group.name.toLowerCase() === 'sample') return 'rgba(224, 242, 254, 0.3)';
    return 'transparent';
  };

  const getDotColor = () => {
    if (isLithology) return '#3b82f6';
    if (isAlteration) return '#f97316';
    if (isGeotechnical) return '#22c55e';
    if (isMineralization) return '#3b82f6';
    if (isInfo) return '#64748b';
    if (group.name.toLowerCase() === 'sample') return '#0ea5e9';
    return group.color;
  };

  const getTextColor = () => {
    if (isLithology) return '#1d4ed8';
    if (isAlteration) return '#c2410c';
    if (isGeotechnical) return '#15803d';
    if (isMineralization) return '#1d4ed8';
    if (isInfo) return '#334155';
    if (group.name.toLowerCase() === 'sample') return '#0369a1';
    return '#7e22ce';
  };

  return (
    <div
      ref={ref}
      style={{
        ...style,
        opacity: isDragging ? 0.5 : 1,
        cursor: 'move',
        width: isInfo ? '100%' : '100%',
        margin: '0.25rem',
      }}
      className="bg-white rounded-lg shadow-sm"
    >
      <div 
        className="px-2 py-0.5 rounded-t-lg flex items-center justify-between"
        style={{ 
          backgroundColor: getHeaderColor()
        }}
      >
        <div className="flex items-center gap-2">
          <FiMove className="text-gray-500" size={14} />
          <div 
            className="w-2 h-2 rounded-full flex-shrink-0" 
            style={{ 
              backgroundColor: getDotColor()
            }}
          />
          <h3 className="text-[15px] font-medium" style={{
            color: getTextColor()
          }}>
            {group.name}
            <span className="ml-1 text-[15px] opacity-70">
              ({group.columns.length})
            </span>
          </h3>
        </div>
      </div>
      <div className="p-0.5" style={{
        backgroundColor: getBackgroundColor()
      }}>
        <div className={`grid ${
          isInfo ? 'grid-cols-2' :
          isLithology ? 'grid-cols-2 lg:grid-cols-2' :
          isAlteration ? 'grid-cols-4 lg:grid-cols-4' :
          isGeotechnical ? 'grid-cols-2 lg:grid-cols-2' :
          isMineralization ? 'grid-cols-2 lg:grid-cols-2' :
          group.name.toLowerCase() === 'sample' ? 'grid-cols-1' :
          'grid-cols-4 lg:grid-cols-5'
        } gap-0.5`}>
          {group.columns.sort((a, b) => {
            if (isLithology) {
              if (a === 'Lith1_GrainSize' && b === 'Lith2               Code') return 1;
              if (a === 'Lith2               Code' && b === 'Lith1_GrainSize') return -1;
            }
            
            if (isMineralization) {
              const mineralizationOrder = [
                'Vein          Type',
                'QzVn %',
                'Bx                   Clast',
                'Bx                  Matrix',
                'QzBx %',
                'Min                    Zone',
                'Pima                   Sample',
                'Comments'
              ];
              return mineralizationOrder.indexOf(a) - mineralizationOrder.indexOf(b);
            }

            if (isAlteration) {
              const alterationOrder = [
                'Sil            Deg',
                'Vuggy%',
                'Pyrite',
                'Alunite',
                'Arg                       Kaolinite',
                'Dickite',
                'Serisite',
                'Oxidation',
                'Gypsum',
                'Carbonate',
                'Chlorite',
                'Epidote',
                'Alteration               Type'
              ];
              return alterationOrder.indexOf(a) - alterationOrder.indexOf(b);
            }

            if (isGeotechnical) {
              const geotechnicalOrder = [
                'recovery_m',
                'rqd',
                'Frac',
                'Mjr_defect               alpha',
                'Weathering',
                'Rock                      Strength',
                'Mjr_defect              type',
              ];
              return geotechnicalOrder.indexOf(a) - geotechnicalOrder.indexOf(b);
            }

            if (isInfo) {
              const infoOrder = [
                'E1',
                'E2',
              ];
              const aIndex = infoOrder.indexOf(a);
              const bIndex = infoOrder.indexOf(b);
              if (aIndex !== -1 && bIndex !== -1) {
                return aIndex - bIndex;
              }
            }
            
            return a.localeCompare(b);
          }).filter(column => column !== 'E2_INPUT').map((column, columnIndex) => (
            <div key={columnIndex} className={`min-w-[100px] bg-white rounded p-0.5 ${isInfo ? 'h-[50px]' : 'h-[45px]'}`}>
              <Label 
                htmlFor={column} 
                className={`text-[13px] ${
                  isAlteration ? 'text-orange-700' :
                  isGeotechnical ? 'text-green-700' :
                  isMineralization ? 'text-blue-700' :
                  isLithology ? 'text-blue-700' :
                  'text-muted-foreground'
                } font-medium mb-0.5 block truncate`}
                title={column}
              >
                {column === 'E1_HEADER' ? 'Dataset' : column}
              </Label>
              {renderFormField(column)}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export const ExcelForm: React.FC = () => {
  const [excelData, setExcelData] = useState<ExcelData | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [currentRowIndex, setCurrentRowIndex] = useState<number>(0);
  const [formData, setFormData] = useState<Record<string, any>>({});
  const [isSaving, setIsSaving] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [__, setExpandedGroups] = useState<Record<string, boolean>>({});
  const [groupOrder, setGroupOrder] = useState<string[]>(() => {
    const saved = localStorage.getItem('groupOrder');
    return saved ? JSON.parse(saved) : [];
  });
  const [lithCodes, setLithCodes] = useState<string[]>([]);
  const [zoneCodes, setZoneCodes] = useState<string[]>([]);
  const [weatheringCodes, setWeatheringCodes] = useState<string[]>([]);
  const [mjrDefectTypes, setMjrDefectTypes] = useState<string[]>([]);
  const [rockStrengthValues, setRockStrengthValues] = useState<string[]>([]);
  const [alterationTypes, setAlterationTypes] = useState<string[]>([]);
  const [minZoneValues, setMinZoneValues] = useState<string[]>([]);

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

  const handleInputChange = (columnName: string, value: string) => {
    if (!excelData) return;

    // Mevcut satırın verilerini güncelle
    const updatedRows = [...excelData.rows];
    const currentRowData = { ...updatedRows[currentRowIndex], [columnName]: value };
    updatedRows[currentRowIndex] = currentRowData;

    // Dataset (E1_HEADER) veya Hole ID değiştiğinde tüm satırlarda güncelle
    if (columnName === 'E1_HEADER' || columnName === 'Hole ID') {
      // Tüm satırlarda güncelle
      updatedRows.forEach((row, index) => {
        updatedRows[index] = {
          ...row,
          [columnName]: value
        };
      });

      // COLLAR verilerini de güncelle
      if (excelData.collarData) {
        const updatedCollarData = {
          ...excelData.collarData,
          [columnName === 'E1_HEADER' ? 'Dataset' : 'Hole ID']: value
        };

        setExcelData({
          ...excelData,
          rows: updatedRows,
          collarData: updatedCollarData
        });
      }
    } else {
      // m_To değiştiğinde ve sonraki satır varsa, sonraki satırın m_From değerini güncelle
      if (columnName === 'm_To' && currentRowIndex < updatedRows.length - 1) {
        updatedRows[currentRowIndex + 1] = {
          ...updatedRows[currentRowIndex + 1],
          m_From: value
        };
      }

      // ExcelData'yı güncelle
      setExcelData({
        ...excelData,
        rows: updatedRows
      });
    }

    // Form verilerini güncelle
    setFormData(currentRowData);
  };

  // COLLAR verilerini güncelleme fonksiyonu
  const handleCollarInputChange = (key: string, value: string) => {
    if (!excelData?.collarData) return;

    const updatedCollarData = {
      ...excelData.collarData,
      [key]: value
    };

    // Dataset veya Hole ID değiştiğinde INFO grubundaki karşılık gelen değerleri de güncelle
    if (key === 'Dataset' || key === 'Hole ID') {
      const updatedRows = excelData.rows.map(row => ({
        ...row,
        [key === 'Dataset' ? 'E1_HEADER' : 'Hole ID']: value
      }));

      setExcelData({
        ...excelData,
        rows: updatedRows,
        collarData: updatedCollarData
      });

      // Mevcut form verilerini de güncelle
      setFormData(updatedRows[currentRowIndex]);
    } else {
      setExcelData({
        ...excelData,
        collarData: updatedCollarData
      });
    }
  };

  const handleNextRow = () => {
    if (!excelData || currentRowIndex >= excelData.rows.length - 1) return;

    const nextIndex = currentRowIndex + 1;
    const updatedRows = [...excelData.rows];
    
    // Mevcut satırın m_To değerini al
    const currentM_To = formatValue(updatedRows[currentRowIndex].m_To);
    
    // Sonraki satırın m_From değerini güncelle
    updatedRows[nextIndex] = {
      ...updatedRows[nextIndex],
      m_From: currentM_To
    };

    // ExcelData'yı güncelle
    setExcelData({
      ...excelData,
      rows: updatedRows
    });

    // Satır indeksini güncelle ve yeni form verilerini ayarla
    setCurrentRowIndex(nextIndex);
    setFormData(updatedRows[nextIndex]);
  };

  const handlePreviousRow = () => {
    if (!excelData || currentRowIndex <= 0) return;

    const prevIndex = currentRowIndex - 1;
    setCurrentRowIndex(prevIndex);
    setFormData(excelData.rows[prevIndex]);
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

  const handleSave = async () => {
    if (!excelData) return;

    setIsSaving(true);
    try {
      // COLLAR verilerini ve mevcut satır verilerini kaydet
      const success = await excelService.saveChanges(currentRowIndex, formData, excelData.collarData as any);
      
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

  // İzin verilen SAMPLE alanları ekleniyor
  const allowedSampleColumns = [
    'sample_this',
    'Sample Number'
  ];

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
        
        // İzin verilen SAMPLE alanları ekleniyor
        if (groupName.toLowerCase().includes('sample')) {
          filteredColumns = filteredColumns.filter(column => 
            allowedSampleColumns.includes(column)
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
        
        // İzin verilen SAMPLE alanları ekleniyor
        if (groupName.toLowerCase().includes('sample')) {
          columns = columns.filter(column => 
            allowedSampleColumns.includes(column)
          );
        }
        
        groupMap.set(groupName, { ...group, name: groupName, columns });
      }
    });

    return Array.from(groupMap.values());
  }, [excelData?.groups]);

  // Excel yüklendiğinde değerleri alalım
  useEffect(() => {
    if (excelData) {
      const loadCodes = async () => {
        try {
          // LithCode değerlerini yükle
          const codes = await excelService.getLithCodesFromDataSheet();
          if (codes && codes.length > 0) {
            setLithCodes(codes);
          } else {
            setLithCodes([
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
            ]);
          }

          // Zone değerlerini yükle
          const zones = await excelService.getZoneCodesFromDataSheet();
          if (zones && zones.length > 0) {
            setZoneCodes(zones);
          }

          // Weathering değerlerini yükle
          const weatherings = await excelService.getWeatheringCodesFromDataSheet();
          if (weatherings && weatherings.length > 0) {
            setWeatheringCodes(weatherings);
          }

          // Mjr_defect type değerlerini yükle
          const defectTypes = await excelService.getMjrDefectTypesFromDataSheet();
          if (defectTypes && defectTypes.length > 0) {
            setMjrDefectTypes(defectTypes);
          }

          // Rock Strength değerlerini yükle
          const strengthValues = await excelService.getRockStrengthFromDataSheet();
          if (strengthValues && strengthValues.length > 0) {
            setRockStrengthValues(strengthValues);
          }

          // Alteration Type değerlerini yükle
          const alterationTypeValues = await excelService.getAlterationTypesFromDataSheet();
          if (alterationTypeValues && alterationTypeValues.length > 0) {
            setAlterationTypes(alterationTypeValues);
          }

          // Min Zone değerlerini yükle
          const minZones = await excelService.getMinZoneFromDataSheet();
          if (minZones && minZones.length > 0) {
            setMinZoneValues(minZones);
          }
        } catch (error) {
          console.error('Kod değerleri yüklenirken hata:', error);
          toast.error('Kod değerleri yüklenemedi');
        }
      };

      loadCodes();
    }
  }, [excelData]);

  // LithCode sütunları için özel dropdown
  const renderLithCodeDropdown = (column: string) => {
    const currentValue = formatValue(formData[column]);
    
    return (
      <div className="relative">
        <Select
          defaultValue={currentValue || "ALL"}
          value={currentValue || "ALL"}
          onValueChange={(value) => handleInputChange(column, value)}
          disabled={isSaving}
        >
          <SelectTrigger className="h-6 w-full text-[14px]">
            <SelectValue placeholder="G" />
          </SelectTrigger>
          <SelectContent>
            {lithCodes.map((code) => (
              <SelectItem key={code} value={code} className="text-[14px]">
                {code}
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
     
      </div>
    );
  };

  // Zone sütunu için özel dropdown
  const renderZoneDropdown = (column: string) => {
    const currentValue = formatValue(formData[column]);
    
    return (
      <div className="relative">
        <Select
          defaultValue={currentValue}
          value={currentValue}
          onValueChange={(value) => handleInputChange(column, value)}
          disabled={isSaving}
        >
          <SelectTrigger className="h-6 w-full text-[14px]">
            <SelectValue placeholder="" />
          </SelectTrigger>
          <SelectContent>
            {zoneCodes.map((code) => (
              <SelectItem key={code} value={code} className="text-[14px]">
                {code}
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
     
      </div>
    );
  };

  // Weathering sütunu için özel dropdown
  const renderWeatheringDropdown = (column: string) => {
    const currentValue = formatValue(formData[column]);
    
    return (
      <div className="relative">
        <Select
          defaultValue={currentValue}
          value={currentValue}
          onValueChange={(value) => handleInputChange(column, value)}
          disabled={isSaving}
        >
          <SelectTrigger className="h-6 w-full text-[14px]">
            <SelectValue placeholder="" />
          </SelectTrigger>
          <SelectContent>
            {weatheringCodes.map((code) => (
              <SelectItem key={code} value={code} className="text-[14px]">
                {code}
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
      
      </div>
    );
  };

  // Mjr_defect type sütunu için özel dropdown
  const renderMjrDefectTypeDropdown = (column: string) => {
    const currentValue = formatValue(formData[column]);
    
    return (
      <div className="relative">
        <Select
          defaultValue={currentValue}
          value={currentValue}
          onValueChange={(value) => handleInputChange(column, value)}
          disabled={isSaving}
        >
          <SelectTrigger className="h-6 w-full text-[14px]">
            <SelectValue placeholder="" />
          </SelectTrigger>
          <SelectContent>
            {mjrDefectTypes.map((type) => (
              <SelectItem key={type} value={type} className="text-[14px]">
                {type}
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>
    );
  };

  // Rock Strength sütunu için özel dropdown
  const renderRockStrengthDropdown = (column: string) => {
    const currentValue = formatValue(formData[column]);
    
    return (
      <div className="relative">
        <Select
          defaultValue={currentValue}
          value={currentValue}
          onValueChange={(value) => handleInputChange(column, value)}
          disabled={isSaving}
        >
          <SelectTrigger className="h-6 w-full text-[14px]">
            <SelectValue placeholder="" />
          </SelectTrigger>
          <SelectContent>
            {rockStrengthValues.map((strength) => (
              <SelectItem key={strength} value={strength} className="text-[14px]">
                {strength}
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>
    );
  };

  // Alteration Type sütunu için özel dropdown
  const renderAlterationTypeDropdown = (column: string) => {
    const currentValue = formatValue(formData[column]);
    
    return (
      <div className="relative">
        <Select
          defaultValue={currentValue}
          value={currentValue}
          onValueChange={(value) => handleInputChange(column, value)}
          disabled={isSaving}
        >
          <SelectTrigger className="h-6 w-full text-[14px]">
            <SelectValue placeholder="" />
          </SelectTrigger>
          <SelectContent>
            {alterationTypes.map((type) => (
              <SelectItem key={type} value={type} className="text-[14px]">
                {type}
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>
    );
  };

  // Min Zone sütunu için özel dropdown
  const renderMinZoneDropdown = (column: string) => {
    const currentValue = formatValue(formData[column]);
    
    return (
      <div className="relative">
        <Select
          defaultValue={currentValue}
          value={currentValue}
          onValueChange={(value) => handleInputChange(column, value)}
          disabled={isSaving}
        >
          <SelectTrigger className="h-6 w-full text-[14px]">
            <SelectValue placeholder="" />
          </SelectTrigger>
          <SelectContent>
            {minZoneValues.map((zone) => (
              <SelectItem key={zone} value={zone} className="text-[14px]">
                {zone}
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>
    );
  };

  const renderFormField = (column: string) => {
    const columnType = excelService.getColumnType(column);
    const dropdownValues = excelService.getDropdownValues(column);
    
    // sample_this için özel dropdown
    if (column === 'sample_this') {
      const currentValue = formatValue(formData[column]);
      return (
        <div className="relative">
          <Select
            defaultValue={currentValue}
            value={currentValue}
            onValueChange={(value) => handleInputChange(column, value)}
            disabled={isSaving}
          >
            <SelectTrigger className="h-6 w-full text-[14px]">
              <SelectValue placeholder="Seçiniz..." />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="Yes" className="text-[14px]">Yes</SelectItem>
              <SelectItem value="No" className="text-[14px]">No</SelectItem>
            </SelectContent>
          </Select>
        </div>
      );
    }

    // E1_HEADER ve E2_INPUT için özel görüntüleme
    if (column === 'E1_HEADER') {
      return (
        <Input
          id={column}
          value={formatValue(formData['E2_INPUT'])}
          disabled={false}
          className="h-6 w-full text-[14px]"
        />
      );
    }

    if (column === 'E2_INPUT') {
      return null;
    }
    
    // LithCode sütunları için özel dropdown
    if (column === 'Lith1               Code' || column === 'Lith2               Code') {
      return renderLithCodeDropdown(column);
    }

    // Zone sütunu için özel dropdown
    if (column.includes('Zone') && column !== 'Min                    Zone') {
      return renderZoneDropdown(column);
    }

    // Min Zone sütunu için özel dropdown
    if (column === 'Min                    Zone') {
      return renderMinZoneDropdown(column);
    }

    // Weathering sütunu için özel dropdown
    if (column === 'Weathering') {
      return renderWeatheringDropdown(column);
    }

    // Mjr_defect type sütunu için özel dropdown
    if (column === 'Mjr_defect              type') {
      return renderMjrDefectTypeDropdown(column);
    }

    // Rock Strength sütunu için özel dropdown
    if (column === 'Rock                      Strength') {
      return renderRockStrengthDropdown(column);
    }

    // Alteration Type sütunu için özel dropdown
    if (column === 'Alteration               Type') {
      return renderAlterationTypeDropdown(column);
    }

    // ALTERATION altındaki alanlar için Rock Strength değerlerini kullanan dropdown
    if (column === 'Arg                       Kaolinite' ||
        column === 'Serisite' ||
        column === 'Dickite' ||
        column === 'Alunite' ||
        column === 'Chlorite' ||
        column === 'Epidote' ||
        column === 'Carbonate' ||
        column === 'Oxidation') {
      return (
        <div className="relative">
          <Select
            defaultValue={formatValue(formData[column])}
            value={formatValue(formData[column])}
            onValueChange={(value) => handleInputChange(column, value)}
            disabled={isSaving}
          >
            <SelectTrigger className="h-6 w-full text-[14px]">
              <SelectValue placeholder="" />
            </SelectTrigger>
            <SelectContent>
              {rockStrengthValues.map((strength) => (
                <SelectItem key={strength} value={strength} className="text-[14px]">
                  {strength}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
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
            <SelectTrigger className="h-6 w-full text-[14px]">
              <SelectValue placeholder="Seçiniz..." />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="default" className="text-[14px]">Seçiniz...</SelectItem>
              {dropdownValues.map((value, index) => (
                <SelectItem key={`${value}-${index}`} value={value} className="text-[14px]">
                  {value}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
          <div className="absolute top-0 right-0 -mr-2 text-[12px] text-muted-foreground">
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
        className="h-6 w-full text-[14px]"
      />
    );
  };

  // Grupları birleştir ve sırala
  const sortedGroups = useMemo(() => {
    if (!excelData?.groups) return [];

    const groups = mergedGroups;
    
    // İlk yükleme kontrolü
    if (groupOrder.length === 0) {
      // INFO grubunu en başa al
      const initialOrder = groups
        .sort((a, b) => {
          if (a.name.toLowerCase() === 'info') return -1;
          if (b.name.toLowerCase() === 'info') return 1;
          return 0;
        })
        .map(group => group.name);
      
      localStorage.setItem('groupOrder', JSON.stringify(initialOrder));
      setGroupOrder(initialOrder);
      return groups;
    }

    // Mevcut sıralamaya göre grupları düzenle
    return [...groups].sort((a, b) => {
      // INFO her zaman en üstte
      if (a.name.toLowerCase() === 'info') return -1;
      if (b.name.toLowerCase() === 'info') return 1;

      const aIndex = groupOrder.indexOf(a.name);
      const bIndex = groupOrder.indexOf(b.name);
      if (aIndex === -1) return 1;
      if (bIndex === -1) return -1;
      return aIndex - bIndex;
    });
  }, [excelData?.groups, groupOrder, mergedGroups]);

  // Grup sırasını güncelle
  const moveBox = useCallback((dragIndex: number, hoverIndex: number) => {
    setGroupOrder(prevOrder => {
      const newOrder = [...prevOrder];
      const [removed] = newOrder.splice(dragIndex, 1);
      newOrder.splice(hoverIndex, 0, removed);
      localStorage.setItem('groupOrder', JSON.stringify(newOrder));
      return newOrder;
    });
  }, []);

  // Grupları özel sıralama ile render et
  const renderGroups = (groups: ExcelGroup[]) => {
    const infoGroup = groups.find(g => g.name.toLowerCase() === 'info');
    const lithologyGroup = groups.find(g => g.name.toLowerCase().includes('lithology'));
    const mineralizationGroup = groups.find(g => g.name.toLowerCase().includes('mineralization'));
    const alterationGroup = groups.find(g => g.name.toLowerCase().includes('alteration'));
    const geotechnicalGroup = groups.find(g => g.name.toLowerCase().includes('geotecnical'));
    const sampleGroup = {
      name: 'SAMPLE',
      color: 'FFFFFF',
      columns: ['sample_this', 'Sample Number'],
      startColumn: 1,
      endColumn: 2
    };

    return (
      <div className="flex flex-col gap-1 w-full">
        {/* INFO ve COLLAR Grubu - En üstte */}
        <div className="flex gap-1 w-full">
          {infoGroup && (
            <div className="w-[25%]">
              <DraggableGroupBox
                key={infoGroup.name}
                group={infoGroup}
                index={groups.indexOf(infoGroup)}
                moveBox={moveBox}
                renderFormField={renderFormField}
                style={{}}
              />
            </div>
          )}
          {excelData?.collarData && (
            <div className="w-[75%]">
              <div className="bg-white rounded-lg shadow-sm m-1">
                <div className="px-2 py-0.5 rounded-t-lg flex items-center justify-between bg-slate-100">
                  <div className="flex items-center gap-2">
                    <div className="w-2 h-2 rounded-full bg-slate-500" />
                    <h3 className="text-xs font-medium text-slate-700">
                      COLLAR
                      <span className="ml-1 text-[15px] opacity-70">
                        ({Object.keys(excelData.collarData).length})
                      </span>
                    </h3>
                  </div>
                </div>
                <div className="p-0.5 bg-slate-50/30">
                  <div className="grid grid-cols-7 gap-0.5">
                    {Object.entries(excelData.collarData).map(([key, value], index) => (
                      <div key={index} className="min-w-[100px] bg-white rounded p-0.5 h-[45px]">
                        <Label
                          htmlFor={key}
                          className="text-[13px] text-gray-600 font-medium mb-0.5 block truncate"
                          title={key}
                        >
                          {key}
                        </Label>
                        <Input
                          id={key}
                          value={formatValue(value)}
                          onChange={(e) => handleCollarInputChange(key, e.target.value)}
                          disabled={isSaving}
                          className="h-6 w-full text-[14px]"
                        />
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* LITHOLOGY ve MINERALIZATION ve GEOTECHNICAL - Orta Sırada */}
        <div className="flex gap-1 w-full">
          {lithologyGroup && (
            <div className="flex flex-col w-[25%]">
              <DraggableGroupBox
                key={lithologyGroup.name}
                group={lithologyGroup}
                index={groups.indexOf(lithologyGroup)}
                moveBox={moveBox}
                renderFormField={renderFormField}
                style={{}}
              />
              {mineralizationGroup && (
                <DraggableGroupBox
                  key={mineralizationGroup.name}
                  group={mineralizationGroup}
                  index={groups.indexOf(mineralizationGroup)}
                  moveBox={moveBox}
                  renderFormField={renderFormField}
                  style={{ margin: '0.25rem 0.25rem 0' }}
                />
              )}
            </div>
          )}
          {alterationGroup && (
            <div className="w-[50%]">
              <DraggableGroupBox
                key={alterationGroup.name}
                group={alterationGroup}
                index={groups.indexOf(alterationGroup)}
                moveBox={moveBox}
                renderFormField={renderFormField}
                style={{}}
              />
            </div>
          )}
          {geotechnicalGroup && (
            <div className="w-[25%] flex flex-col">
              <DraggableGroupBox
                key={geotechnicalGroup.name}
                group={geotechnicalGroup}
                index={groups.indexOf(geotechnicalGroup)}
                moveBox={moveBox}
                renderFormField={renderFormField}
                style={{}}
              />
              <div className="w-full mt-1">
                <DraggableGroupBox
                  key={sampleGroup.name}
                  group={sampleGroup}
                  index={-1}
                  moveBox={moveBox}
                  renderFormField={renderFormField}
                  style={{}}
                />
              </div>
            </div>
          )}
        </div>
      </div>
    );
  };

  return (
    <DndProvider backend={HTML5Backend}>
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
                    <span className="text-[17px] font-bold">Satır {currentRowIndex + 1} / {excelData.rows.length}</span>
                  </div>
                  <Button 
                    onClick={handleSave} 
                    size="sm"
                    disabled={isSaving}
                    className="h-6 px-2 flex items-center gap-1 text-[14px]"
                  >
                    <FiSave size={12} />
                    {isSaving ? 'Kaydediliyor...' : 'Kaydet'}
                  </Button>
                </div>
              </div>
              <div className="overflow-auto h-[calc(100vh-3rem)] p-2">
                {renderGroups(sortedGroups)}
              </div>
            </div>
            
            {/* Sayfalama Butonları - Sağ Alt Sabit */}
            <div className="fixed bottom-2 right-2 flex items-center gap-2 bg-background/80 backdrop-blur-sm p-2 rounded-md shadow-lg z-50">
              <Button
                variant="outline"
                size="sm"
                onClick={() => {
                  const newIndex = Math.max(0, currentRowIndex - 10);
                  setCurrentRowIndex(newIndex);
                  setFormData(excelData.rows[newIndex]);
                }}
                disabled={currentRowIndex === 0 || isSaving}
                className="h-8 px-2 flex items-center gap-1"
              >
                <FiChevronLeft size={14} />
                <FiChevronLeft size={14} className="-ml-1" />
              </Button>
              <Button
                variant="outline"
                size="sm"
                onClick={handlePreviousRow}
                disabled={currentRowIndex === 0 || isSaving}
                className="h-8 w-8 p-0"
              >
                <FiChevronLeft size={14} />
              </Button>
              <span className="text-[16px] font-medium px-2 min-w-[100px] text-center">
                {currentRowIndex + 1} / {excelData.rows.length}
              </span>
              <Button
                variant="outline"
                size="sm"
                onClick={handleNextRow}
                disabled={currentRowIndex === excelData.rows.length - 1 || isSaving}
                className="h-8 w-8 p-0"
              >
                <FiChevronRight size={14} />
              </Button>
              <Button
                variant="outline"
                size="sm"
                onClick={() => {
                  const newIndex = Math.min(excelData.rows.length - 1, currentRowIndex + 10);
                  setCurrentRowIndex(newIndex);
                  setFormData(excelData.rows[newIndex]);
                }}
                disabled={currentRowIndex === excelData.rows.length - 1 || isSaving}
                className="h-8 px-2 flex items-center gap-1"
              >
                <FiChevronRight size={14} />
                <FiChevronRight size={14} className="-ml-1" />
              </Button>
            </div>
          </div>
        )}
      </div>
    </DndProvider>
  );
}; 