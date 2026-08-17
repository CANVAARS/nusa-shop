/* =====================================================================
   TOKO BERSAMA — BACKEND APPS SCRIPT (Code.gs)
   Perbaikan:
   1. URL gambar Drive yang salah (dulu: "https://googleusercontent.com"+id -> rusak,
      sekarang pakai format thumbnail Drive yang benar dan bisa tampil di <img>)
   2. Aksi tambah seller/admin di frontend mengirim "add_user" tapi backend
      hanya mengenali "admin_add_seller" -> disamakan jadi "admin_add_seller"
      (diperbaiki di sisi login.html)
   3. Verifikasi login disempurnakan (trim, cek sheet & header)
   4. FITUR BARU: Komentar + Rating bintang (1-5) per produk
      - doPost action "post_komentar"  -> kirim komentar & rating (publik, tanpa login)
      - doGet  action "get_komentar"   -> ambil daftar komentar + rata2 rating 1 produk
      - doGet  action "get_ringkasan_rating" -> ambil rata2 rating SEMUA produk sekaligus
   ===================================================================== */

function doPost(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    var data = JSON.parse(e.postData.contents);
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // =============================================================
    // ACTION: LOGIN (ADMIN & SELLER)
    // =============================================================
    if (data.action === "login") {
      var userStatus = verifyUser(spreadsheet, data.username, data.passwordHash);
      if (userStatus.authenticated) {
        return res({ status: "success", role: userStatus.role, message: "Login Berhasil!" }, headers);
      } else {
        return res({ status: "error", message: "Username atau Password salah!" }, headers);
      }
    }

    // =============================================================
    // ACTION: KOMENTAR + RATING (PUBLIK — TIDAK PERLU LOGIN)
    // =============================================================
    if (data.action === "post_komentar") {
      var produkId = String(data.produkId || "").trim();
      var namaKomentator = String(data.nama || "").trim();
      var isiKomentar = String(data.komentar || "").trim();
      var rating = parseInt(data.rating, 10);

      if (!produkId) {
        return res({ status: "error", message: "Produk tidak valid." }, headers);
      }
      if (!namaKomentator) {
        return res({ status: "error", message: "Nama wajib diisi." }, headers);
      }
      if (isNaN(rating) || rating < 1 || rating > 5) {
        return res({ status: "error", message: "Rating harus antara 1 sampai 5 bintang." }, headers);
      }

      var komentarSheet = getOrCreateKomentarSheet(spreadsheet);
      var komentarId = "CMT-" + new Date().getTime();
      var waktu = new Date();

      komentarSheet.appendRow([
        komentarId,
        produkId,
        namaKomentator,
        isiKomentar,
        rating,
        waktu
      ]);

      return res({ status: "success", message: "Terima kasih! Komentar & rating kamu sudah tersimpan." }, headers);
    }

    // =============================================================
    // VERIFIKASI KEAMANAN UNTUK AKSI YANG BUTUH LOGIN (di bawah ini)
    // =============================================================
    var userCheck = verifyUser(spreadsheet, data.username, data.passwordHash);
    if (!userCheck.authenticated) {
      return res({ status: "error", message: "Akses ditolak! Sesi tidak valid, silakan login ulang." }, headers);
    }

    // =============================================================
    // ACTION: POST PRODUK (SELLER & ADMIN)
    // =============================================================
    if (data.action === "post_produk") {
      if (!data.image || !data.nama || !data.harga || !data.kontak) {
        return res({ status: "error", message: "Data produk tidak lengkap." }, headers);
      }

      var folderId = "1LdAqavSYRNGQWZwcHN45u2mJq1gRUnJn";
      var folder = DriveApp.getFolderById(folderId);

      var contentType = data.image.match(/data:(.*?);/);
      var mimeType = contentType ? contentType[1] : "image/jpeg";
      var base64Data = data.image.replace(/^data:image\/\w+;base64,/, "");
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, data.imageName || "produk.jpg");

      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      // PERBAIKAN: format URL gambar Drive yang benar agar tampil di <img>
      var imageUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";

      var productSheet = spreadsheet.getSheetByName("Produk");
      var productId = "PRD-" + new Date().getTime();

      productSheet.appendRow([
        productId,
        data.nama,
        data.harga,
        data.keterangan,
        data.kontak,
        imageUrl,
        data.username.toLowerCase().trim()
      ]);

      return res({ status: "success", message: "Produk berhasil diterbitkan!", productId: productId, imageUrl: imageUrl }, headers);
    }

    // =============================================================
    // ACTION: FITUR KHUSUS ADMIN (TAMBAH SELLER/ADMIN BARU)
    // =============================================================
    if (userCheck.role !== "admin") {
      return res({ status: "error", message: "Anda tidak memiliki hak akses Admin!" }, headers);
    }

    if (data.action === "admin_add_seller") {
      if (!data.newUsername || !data.newPasswordHash || !data.newRole) {
        return res({ status: "error", message: "Data akun baru tidak lengkap." }, headers);
      }

      var sellerSheet = spreadsheet.getSheetByName("Sellers");
      var rows = sellerSheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).toLowerCase().trim() === String(data.newUsername).toLowerCase().trim()) {
          return res({ status: "error", message: "Username sudah terdaftar!" }, headers);
        }
      }

      sellerSheet.appendRow([
        String(data.newUsername).toLowerCase().trim(),
        String(data.newPasswordHash).toLowerCase().trim(),
        String(data.newRole).toLowerCase().trim()
      ]);
      return res({ status: "success", message: "User/Seller baru berhasil didaftarkan!" }, headers);
    }

    return res({ status: "error", message: "Aksi tidak dikenali." }, headers);

  } catch (error) {
    return res({ status: "error", message: error.toString() }, headers);
  }
}

function doGet(e) {
  var headers = { "Access-Control-Allow-Origin": "*" };

  try {
    var action = e.parameter.action;
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // =============================================================
    // ACTION: AMBIL KOMENTAR + RATING SATU PRODUK
    // =============================================================
    if (action === "get_komentar") {
      var produkId = String(e.parameter.produkId || "").trim();
      if (!produkId) {
        return res({ status: "error", message: "produkId wajib diisi." }, headers);
      }

      var komentarSheet = getOrCreateKomentarSheet(spreadsheet);
      var rows = komentarSheet.getDataRange().getValues();
      var daftarKomentar = [];
      var totalRating = 0;
      var jumlah = 0;

      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][1]).trim() === produkId) {
          var ratingVal = Number(rows[i][4]) || 0;
          daftarKomentar.push({
            nama: String(rows[i][2]),
            komentar: String(rows[i][3]),
            rating: ratingVal,
            waktu: rows[i][5] instanceof Date ? rows[i][5].toISOString() : String(rows[i][5])
          });
          totalRating += ratingVal;
          jumlah++;
        }
      }

      daftarKomentar.reverse(); // terbaru dulu

      return res({
        status: "success",
        produkId: produkId,
        komentar: daftarKomentar,
        rataRata: jumlah > 0 ? Math.round((totalRating / jumlah) * 10) / 10 : 0,
        jumlah: jumlah
      }, headers);
    }

    // =============================================================
    // ACTION: RINGKASAN RATING SEMUA PRODUK SEKALIGUS
    // =============================================================
    if (action === "get_ringkasan_rating") {
      var komentarSheet2 = getOrCreateKomentarSheet(spreadsheet);
      var rows2 = komentarSheet2.getDataRange().getValues();
      var ringkasan = {};

      for (var j = 1; j < rows2.length; j++) {
        var pid = String(rows2[j][1]).trim();
        if (!pid) continue;
        var r = Number(rows2[j][4]) || 0;
        if (!ringkasan[pid]) ringkasan[pid] = { total: 0, jumlah: 0 };
        ringkasan[pid].total += r;
        ringkasan[pid].jumlah += 1;
      }

      var hasil = {};
      for (var key in ringkasan) {
        hasil[key] = {
          rataRata: Math.round((ringkasan[key].total / ringkasan[key].jumlah) * 10) / 10,
          jumlah: ringkasan[key].jumlah
        };
      }

      return res({ status: "success", ringkasan: hasil }, headers);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sistem mendeteksi GET tanpa action yang dikenali. Gunakan POST text/plain murni untuk login/post produk." }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders(headers);

  } catch (error) {
    return res({ status: "error", message: error.toString() }, headers);
  }
}

/* ---------------- HELPER: SHEET KOMENTAR ---------------- */
function getOrCreateKomentarSheet(spreadsheet) {
  var sheet = spreadsheet.getSheetByName("Komentar");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("Komentar");
    sheet.appendRow(["id", "produkId", "nama", "komentar", "rating", "waktu"]);
  }
  return sheet;
}

/* ---------------- VERIFIKASI LOGIN ---------------- */
function verifyUser(spreadsheet, username, passwordHash) {
  if (!username || !passwordHash) {
    return { authenticated: false, role: null };
  }

  var sellerSheet = spreadsheet.getSheetByName("Sellers");
  if (!sellerSheet) {
    return { authenticated: false, role: null };
  }

  var sellerData = sellerSheet.getDataRange().getValues();

  var u = String(username).toLowerCase().trim();
  var p = String(passwordHash).toLowerCase().trim();

  for (var i = 1; i < sellerData.length; i++) {
    var dbUsername = String(sellerData[i][0]).toLowerCase().trim();
    var dbPasswordHash = String(sellerData[i][1]).toLowerCase().trim();
    var dbRole = String(sellerData[i][2]).toLowerCase().trim();

    if (dbUsername === u && dbPasswordHash === p) {
      return { authenticated: true, role: dbRole };
    }
  }
  return { authenticated: false, role: null };
}

function res(outputObject, headers) {
  return ContentService.createTextOutput(JSON.stringify(outputObject))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders(headers);
}

/* =====================================================================
   JALANKAN SEKALI SAJA SECARA MANUAL DARI EDITOR APPS SCRIPT
   untuk membuat hash password admin pertama kamu.
   1. Ganti nilai "GANTI_PASSWORD_ADMIN_DISINI" di bawah.
   2. Pilih fungsi "buatHashPasswordAwal" lalu klik Run.
   3. Buka View > Logs (Ctrl+Enter) untuk melihat hasil hash-nya.
   4. Salin hash tsb ke sheet "Sellers": kolom A=username, B=hash, C=admin
   ===================================================================== */
function buatHashPasswordAwal() {
  var password = "GANTI_PASSWORD_ADMIN_DISINI";
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  var hexHash = rawHash.map(function (byte) {
    var v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
  Logger.log("Hash untuk password '" + password + "' adalah:");
  Logger.log(hexHash);
}
