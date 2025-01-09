const express = require('express');
const cors = require('cors');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const os = require('os');

const app = express();

// Multer yapılandırması
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    // Geçici klasör oluştur
    const tempDir = path.join(os.tmpdir(), 'excel-updates');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    cb(null, tempDir);
  },
  filename: function (req, file, cb) {
    // Orijinal dosya adını koru
    cb(null, file.originalname);
  }
});

const upload = multer({ storage: storage });

app.use(cors());
app.use(express.json());

// Excel dosyasını güncelleme endpoint'i
app.post('/api/update-excel', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Dosya eksik' });
    }

    const fileName = req.body.fileName;
    let filePath = req.body.filePath;

    // Eğer filePath boşsa, Downloads klasörüne kaydet
    if (!filePath) {
      const homeDir = os.homedir();
      const downloadsDir = path.join(homeDir, 'Downloads');
      filePath = path.join(downloadsDir, fileName);
    }

    console.log('Dosya güncelleniyor:', {
      geçiciDosya: req.file.path,
      hedefDosya: filePath
    });

    try {
      // Dosyayı güncelle - doğrudan kopyalama yerine içeriği oku ve yaz
      const fileContent = fs.readFileSync(req.file.path);
      fs.writeFileSync(filePath, fileContent);
      
      // Geçici dosyayı sil
      fs.unlinkSync(req.file.path);

      console.log('Dosya başarıyla güncellendi:', filePath);
      res.json({ 
        success: true,
        message: 'Dosya güncellendi',
        path: filePath
      });
    } catch (error) {
      console.error('Dosya işlem hatası:', error);
      res.status(500).json({ 
        error: 'Dosya işlem hatası',
        details: error.message
      });
    }
  } catch (error) {
    console.error('Genel hata:', error);
    res.status(500).json({ 
      error: 'Sunucu hatası',
      details: error.message
    });
  }
});

const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`Server ${PORT} portunda çalışıyor`);
}); 