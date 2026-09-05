/**
 * ====================================================================
 * PORTAL TRAH SILSILAH & KOMUNIKASI KELUARGA BESAR (FULL COMPREHENSIVE BUILD)
 * Backend Google Apps Script & Database Spreadsheet Handler
 * ====================================================================
 */

var MANUAL_SPREADSHEET_ID = "";

function getSS() {
  if (MANUAL_SPREADSHEET_ID && MANUAL_SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(MANUAL_SPREADSHEET_ID.trim());
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Spreadsheet belum terhubung! Buka Google Sheets > Ekstensi > Apps Script atau isi MANUAL_SPREADSHEET_ID di Code.gs");
  }
  return ss;
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal Silsilah & Komunikasi Keluarga')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  return doGet();
}

/**
 * Inisialisasi Seluruh Sheet & Struktur Data Awal
 */
function setupDatabase() {
  var ss = getSS();

  var schemas = {
    "Users": [
      ["id", "username", "password", "nama", "role"],
      ["USR-001", "admin", "keluarga123", "Bapak Yudho Purwoko", "Admin"],
      ["USR-002", "keluarga", "123456", "Anggota Keluarga", "Anggota"]
    ],
    "Anggota": [
      ["id", "nama", "gender", "generasi", "ayah", "anak_ke", "pasangan", "id_ortu", "tgl_lahir", "alamat_ringkas", "alamat_lengkap", "pekerjaan", "keistimewaan", "bisnis", "riwayat_ringkas", "riwayat_lengkap", "telepon", "foto_url"],
      ["MBR-001", "Bapak Tohari", "L", 1, "-", 1, "Ibu Siti Aminah", "", "1940-05-12", "Solo, Jawa Tengah", "Jl. Slamet Riyadi No. 45, Laweyan, Surakarta", "Pendidik & Tokoh Adat", "Hafal silsilah 7 turunan & pengayom trah keluarga", "Perkebunan & Toko Buku Pusaka", "Sesepuh keluarga yang mendidik dengan kesederhanaan.", "Lahir di era perintis kemerdekaan. Mendedikasikan hidupnya untuk pendidikan rakyat dan memelihara kerukunan antar generasi trah keluarga besar.", "081234567801", "https://images.unsplash.com/photo-1506794778202-cad84cf45f1d?w=200"],
      ["MBR-002", "Ibu Siti Aminah", "P", 1, "-", 1, "Bapak Tohari", "", "1945-08-20", "Solo, Jawa Tengah", "Jl. Slamet Riyadi No. 45, Laweyan, Surakarta", "Wiraswasta", "Memasak kuliner khas legendaris & pemersatu cucu", "Batik Tulis Pusaka", "Penyayang, selalu menjadi pusat keceriaan di setiap reuni trah.", "Mendampingi keluarga selama puluhan tahun, mengayomi anak cucu dan aktif melestarikan tradisi luhur keluarga.", "081234567802", "https://images.unsplash.com/photo-1544005313-94ddf0286df2?w=200"],
      ["MBR-003", "Yudho Purwoko", "L", 2, "Bapak Tohari", 5, "Istri Tercinta", "MBR-001", "1971-03-09", "Depok, Jawa Barat", "Komplek Pesona Kahuripan Blok C, Depok", "Head of Finance & Konsultan Bisnis", "Visioner, teliti tata kelola dan perencana masa depan", "Konsultan Manajemen Bisnis & Infrastruktur", "Anak ke-5 yang aktif merajut silaturahmi seluruh trah.", "Lulusan Psikologi yang berpengalaman panjang di manajemen korporasi, telekomunikasi, dan tata kelola organisasi. Selalu menjadi inisiator digitalisasi keluarga.", "081198765432", "https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=200"],
      ["MBR-004", "Tikmung Kicik", "P", 3, "Yudho Purwoko", 1, "-", "MBR-003", "2008-07-14", "Depok, Jawa Barat", "Komplek Pesona Kahuripan Blok C, Depok", "Pelajar / Mahasiswa", "Kreatif, melek teknologi dan jago desain visual digital", "Studio Desain Kreatif", "Generasi muda pembawa keceriaan dan inovasi keluarga.", "Generasi cucu yang aktif membantu urusan teknologi informasi, pembuatan konten dokumentasi, dan visual grafis setiap agenda trah.", "081299887766", "https://images.unsplash.com/photo-1494790108377-be9c29b29330?w=200"]
    ],
    "StatusFeed": [
      ["id", "id_anggota", "nama_penulis", "konten", "kategori", "suka_count", "tanggal"],
      ["F-001", "MBR-003", "Yudho Purwoko", "Alhamdulillah arisan putaran lalu lancar berkah. Semoga trah kita selalu guyub rukun!", "Kabar", 15, "2026-05-15 14:00"],
      ["F-002", "MBR-004", "Tikmung Kicik", "Humor sore: Kenapa kakek kalau jalan santai suka bawa tongkat? Soalnya kalau bawa remote TV diperebutkan cucu! Haha", "Humor", 24, "2026-05-16 17:30"]
    ],
    "KomentarStatus": [
      ["id", "id_feed", "nama_komentator", "komentar", "tanggal"],
      ["KMT-001", "F-001", "Bapak Tohari", "Aamiin ya Rabbal alamin, berkah kagem sedoyo.", "2026-05-15 15:10"],
      ["KMT-002", "F-002", "Yudho Purwoko", "Bisa aja Tikmung haha!", "2026-05-16 18:05"]
    ],
    "Kegiatan": [
      ["id", "nama_acara", "tipe", "tanggal_waktu", "lokasi", "tuan_rumah", "deskripsi", "biaya"],
      ["EVT-01", "Arisan Bulanan & Halal Bihalal", "Arisan", "2026-06-07 10:00", "Kediaman Mas Yudho (Depok)", "Keluarga Mas Yudho", "Kocokan arisan putaran ke-6 sekaligus syukuran silaturahmi.", "100000"],
      ["EVT-02", "Khataman & Pengajian Rutin Trah", "Pengajian", "2026-06-21 19:30", "Grup Online Zoom & Kediaman Solo", "Bapak Tohari", "Doa bersama keselamatan sesepuh serta kesuksesan anak cucu.", "0"]
    ],
    "PendaftaranKegiatan": [
      ["id", "id_kegiatan", "nama_peserta", "status_kehadiran", "jumlah_keluarga", "catatan", "tanggal_daftar"],
      ["RSVP-01", "EVT-01", "Bapak Tohari", "Hadir", 2, "Insya Allah hadir membawa oleh-oleh Solo", "2026-05-18"]
    ],
    "GrupChat": [
      ["id", "channel", "nama_pengirim", "pesan", "waktu"],
      ["C-01", "Umum", "Bapak Tohari", "Assalamu'alaikum anak cucu sekalian, semoga sehat semua.", "08:15"],
      ["C-02", "Umum", "Yudho Purwoko", "Wa'alaikumussalam Bapak, aamiin ya rabbal alamin.", "08:20"]
    ],
    "Galeri": [
      ["id", "judul", "kategori", "foto_url", "pengunggah", "tanggal"],
      ["GAL-01", "Silaturahmi Idul Fitri Trah Solo", "Momen Lebaran", "https://images.unsplash.com/photo-1511895426328-dc8714191300?w=600", "Yudho Purwoko", "2026-04-10"],
      ["GAL-02", "Wisuda & Syukuran Generasi Muda", "Pendidikan", "https://images.unsplash.com/photo-1523050854058-8df90110c9f1?w=600", "Tikmung Kicik", "2026-05-02"]
    ],
    "Komunitas": [
      ["id", "nama_komunitas", "deskripsi", "ketua", "jumlah_anggota"],
      ["KOM-01", "Komunitas Sepeda & Gowes Sehat Trah", "Kegiatan bersepeda santai mingguan keluarga besar", "Yudho Purwoko", 12],
      ["KOM-02", "Forum Wirausaha & Bisnis Keluarga", "Kolaborasi usaha, pemasaran produk, dan investasi antar saudara", "Bapak Tohari", 8]
    ]
  };

  for (var sheetName in schemas) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var data = schemas[sheetName];
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      sheet.getRange(1, 1, 1, data[0].length).setBackground("#1b4332").setFontColor("#ffffff").setFontWeight("bold");
    }
  }
  return { status: "success", message: "Database dan seluruh tabel berhasil dibuat!" };
}

function cleanData(values) {
  if (!values || values.length <= 1) return [];
  var headers = values[0];
  var rows = values.slice(1);
  return rows.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      var val = row[index];
      if (val instanceof Date) {
        var y = val.getFullYear();
        var m = ("0" + (val.getMonth() + 1)).slice(-2);
        var d = ("0" + val.getDate()).slice(-2);
        val = y + "-" + m + "-" + d;
      }
      obj[header] = (val !== null && val !== undefined) ? String(val).trim() : "";
    });
    return obj;
  });
}

function getAllData() {
  try {
    var ss = getSS();
    var getSheetRows = function(name) {
      var sheet = ss.getSheetByName(name);
      if (!sheet) return [];
      var vals = sheet.getDataRange().getValues();
      return cleanData(vals);
    };

    return {
      status: "success",
      users: getSheetRows("Users"),
      anggota: getSheetRows("Anggota"),
      statusFeed: getSheetRows("StatusFeed"),
      komentarStatus: getSheetRows("KomentarStatus"),
      kegiatan: getSheetRows("Kegiatan"),
      pendaftaran: getSheetRows("PendaftaranKegiatan"),
      grupChat: getSheetRows("GrupChat"),
      galeri: getSheetRows("Galeri"),
      komunitas: getSheetRows("Komunitas")
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function loginUser(username, password) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Users");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("Users");
    }
    var users = cleanData(sheet.getDataRange().getValues());
    var match = users.find(function(u) {
      return u.username.toLowerCase() === String(username).toLowerCase().trim() &&
             u.password === String(password).trim();
    });

    if (match) {
      return {
        status: "success",
        user: { id: match.id, username: match.username, nama: match.nama, role: match.role }
      };
    } else {
      return { status: "error", message: "Username atau password tidak cocok!" };
    }
  } catch (err) {
    return { status: "error", message: "Gagal memproses autentikasi: " + err.toString() };
  }
}

function uploadFileToDrive(base64Data, fileName, mimeType) {
  try {
    var folderName = "Foto Silsilah Keluarga";
    var folders = DriveApp.getFoldersByName(folderName);
    var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    var decoded = Utilities.base64Decode(base64Data.split(',')[1] || base64Data);
    var blob = Utilities.newBlob(decoded, mimeType || "image/jpeg", fileName || ("Foto_" + new Date().getTime() + ".jpg"));
    
    var file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var fileUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    return { status: "success", fileUrl: fileUrl };
  } catch (err) {
    return { status: "error", message: "Gagal upload ke Drive: " + err.toString() };
  }
}

function saveAnggota(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Anggota");
    var id = "MBR-" + ("000" + sheet.getLastRow()).slice(-3);
    sheet.appendRow([
      id, data.nama, data.gender, data.generasi,
      data.ayah || "-", data.anak_ke || "-", data.pasangan || "-",
      data.id_ortu || "", data.tgl_lahir, data.alamat_ringkas,
      data.alamat_lengkap, data.pekerjaan, data.keistimewaan,
      data.bisnis || "-", data.riwayat_ringkas, data.riwayat_lengkap,
      data.telepon || "-", data.foto_url || "https://images.unsplash.com/photo-1535713875002-d1d0cf377fde?w=200"
    ]);
    return { status: "success", id: id };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function updateAnggota(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Anggota");
    var values = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] == data.id) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) return { status: "error", message: "Anggota tidak ditemukan" };

    sheet.getRange(rowIndex, 2, 1, 17).setValues([[
      data.nama, data.gender, data.generasi,
      data.ayah || "-", data.anak_ke || "-", data.pasangan || "-",
      data.id_ortu || "", data.tgl_lahir, data.alamat_ringkas,
      data.alamat_lengkap, data.pekerjaan, data.keistimewaan,
      data.bisnis || "-", data.riwayat_ringkas, data.riwayat_lengkap,
      data.telepon || "-", data.foto_url || values[rowIndex - 1][17]
    ]]);
    return { status: "success" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function updateFotoProfil(id, foto_url) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Anggota");
    var values = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] == id) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) return { status: "error", message: "ID anggota tidak ditemukan" };

    sheet.getRange(rowIndex, 18).setValue(foto_url);
    return { status: "success" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function saveKomentar(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("KomentarStatus");
    if (!sheet) {
      sheet = ss.insertSheet("KomentarStatus");
      sheet.appendRow(["id", "id_feed", "nama_komentator", "komentar", "tanggal"]);
      sheet.getRange(1, 1, 1, 5).setBackground("#1b4332").setFontColor("#ffffff").setFontWeight("bold");
    }
    var id = "KMT-" + ("000" + sheet.getLastRow()).slice(-3);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd HH:mm");
    sheet.appendRow([id, data.id_feed, data.nama_komentator, data.komentar, now]);
    return { status: "success", id: id, tanggal: now };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function saveFeed(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("StatusFeed");
    var id = "F-" + ("000" + sheet.getLastRow()).slice(-3);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd HH:mm");
    sheet.appendRow([id, data.id_anggota || "USR", data.nama_penulis, data.konten, data.kategori || "Kabar", 0, now]);
    return { status: "success", id: id };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function saveKegiatan(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Kegiatan");
    var id = "EVT-" + ("00" + sheet.getLastRow()).slice(-2);
    sheet.appendRow([id, data.nama_acara, data.tipe, data.tanggal_waktu, data.lokasi, data.tuan_rumah, data.deskripsi, data.biaya || "0"]);
    return { status: "success", id: id };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function savePendaftaran(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("PendaftaranKegiatan");
    var id = "RSVP-" + ("000" + sheet.getLastRow()).slice(-3);
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
    sheet.appendRow([id, data.id_kegiatan, data.nama_peserta, data.status_kehadiran, data.jumlah_keluarga, data.catatan || "-", today]);
    return { status: "success", id: id };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function saveChat(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("GrupChat");
    var id = "C-" + ("000" + sheet.getLastRow()).slice(-3);
    var timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "HH:mm");
    sheet.appendRow([id, data.channel || "Umum", data.nama_pengirim, data.pesan, timeStr]);
    return { status: "success", id: id };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function saveGaleri(data) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Galeri");
    var id = "GAL-" + ("00" + sheet.getLastRow()).slice(-2);
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
    sheet.appendRow([id, data.judul, data.kategori, data.foto_url, data.pengunggah, today]);
    return { status: "success", id: id };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}