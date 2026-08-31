// ═══════════════════════════════════════════════════════════════════
//  CANYAP — E-FATURA + STOK + DASHBOARD (Birleşik v13) — DÜZELTİLMİŞ
//  Hesap: fincanlaryapi@gmail.com
//  Tek dosya: 19t4MsvudC8X7knZ_dymBm5fghcbZcpAMwOmUXZxDPPQ
//
//  ── BU SÜRÜMDE YAPILAN DÜZELTMELER ──
//  1) fiyatYaz(): FATURA_TARIHI sütununu yazmadan önce düz metin ("@")
//     formatına zorlayan satır geri eklendi. Eksik olduğu için Google
//     Sheets "31/10/2026" gibi bir metni otomatik olarak gerçek bir
//     tarih hücresine çeviriyordu; biz de onu okurken JS'in çirkin
//     "Sat Oct 31 2026 00:00:00 GMT+0300 (...)" halini görüyorduk.
//  2) isleMail(): faturaParseEt'e KALICI link (edmLink) yerine yanlışlıkla
//     GEÇİCİ/kısa ömürlü link (htmlLink, "tmp-ef" içeren) gönderiliyordu.
//     Bu yüzden özellikle yeni faturaların "Görüntüle" linki kısa süre
//     sonra 404 veriyordu. edmLink gönderilecek şekilde düzeltildi.
//  3) eskiEdmGuncelle(): mantığı ters kurulmuştu — zaten bozuk (tmp-ef)
//     linkleri "yeni format" sayıp dokunmuyor, sağlam linkleri ise
//     bozuyordu. Bu fonksiyon KALDIRILDI, yerine linkVeTarihOnar()
//     eklendi (aşağıda, en altta) — mevcut bozuk kayıtları TEK SEFERLİK
//     onarmak için kullanılır, çalıştırdıktan sonra bir daha çalıştırmanıza
//     gerek yok.
//  4) ★ YENİ (bu düzenleme): FATURAFIYAT sayfasına MIKTAR kolonu eklendi.
//     faturaParseEt() zaten fatura HTML'inden miktarı parse ediyordu, ama
//     fiyatYaz() bu alanı sayfaya yazarken atlıyordu — sonuç olarak ERP
//     tarafındaki "Bekleyen Alış Faturaları" onay ekranında miktar hep 0
//     geliyordu. miktarKolonuEkle() mevcut sayfaya (bir kereye mahsus,
//     zararsız şekilde tekrar tekrar) MIKTAR kolonunu ekler; fiyatYaz()
//     artık bu kolona yazıyor ve FATURA_TARIHI kolonunu sabit "10" yerine
//     başlıktan dinamik buluyor (kolon numarası MIKTAR eklenince kaydığı
//     için hardcoded index artık yanlış olurdu).
// ═══════════════════════════════════════════════════════════════════

// ── TEK SABİT ──
var SHEET_ID = "19t4MsvudC8X7knZ_dymBm5fghcbZcpAMwOmUXZxDPPQ";

// E-Fatura ayarları
var SHEET_FIYAT   = "FATURAFIYAT";
var SHEET_LOG     = "FATURA_LOG";
var MAIL_QUERY_EF = 'subject:"e-Faturanız var" is:unread';
var EDM_BASE      = "https://view.edmbilisim.com.tr";
var PDF_FOLDER    = "FATURALAR";

var FATURA_FILTRE = []; // boş = tüm tedarikçilerin e-faturaları kabul edilir (önceden sadece "CNY" ön ekli Canyap faturaları)
var NAKLIYE_KW    = ["NAKLİYE","NAKLIYE","PALET","TAŞIMA","TASIMA","SEVK","KARGO","TRANSPORT"];
var NAKLIYE_KODLAR  = [
  "999997000101","331857000011","331857000012",
  "500112174100","468887000003","810103001356"
];
var SERAMIK_KW = [
  "X120","X60","X90","X45","X75","X50","X30","X20","X33","X66",
  "61X61","60X5","42X42","42,5","60,5"
];

// Stok ayarları — filename:Canyap kaldırıldı, filtre kod içinde
var MAIL_QUERY_STOK = "in:inbox has:attachment newer_than:7d";

const STOK_SHEETS = {
  notlar:    "Notlar",
  etiketler: "Etiketler",
  urunEt:    "UrunEtiketleri",
  markalar:  "Markalar",
  ayarlar:   "Ayarlar",
  stoklar:   "Stoklar",
  log:       "Log",
};

// ═══════════════════════════════════════════════════════════════════
//  E-FATURA
// ═══════════════════════════════════════════════════════════════════

function efaturaOku() {
  var baslangic = Date.now();
  var SURE_SINIRI_MS = 4.5 * 60 * 1000; // 6dk'lık sert Apps Script sınırına çarpmadan nazikçe dur

  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var shFiy = getOrCreateSheet(ss, SHEET_FIYAT, [
    "STOK_KODU","STOK_ADI","MIKTAR","BIRIM_FIYAT","ISKONTO","NET_FIYAT",
    "NAKLIYE_PAYI","MALIYET_FIYAT","KDV_ORANI",
    "FATURA_NO","FATURA_TARIHI","TEDARIKCI",
    "FATURA_LINK","EDM_LINK","ISLEM_ZAMANI"
  ]);
  miktarKolonuEkle(shFiy);
  var shLog = getOrCreateSheet(ss, SHEET_LOG,
    ["FATURA_NO","GONDEREN","TARIH","DURUM","DETAY","ISLEM_ZAMANI"]);

  var threads = GmailApp.search(MAIL_QUERY_EF, 0, 200);
  if (threads.length === 0) { Logger.log("Yeni e-fatura yok."); return; }
  Logger.log(threads.length + " e-fatura maili bulundu.");

  var durduruldu = false;
  threads.forEach(function(thread) {
    if (durduruldu) return;
    thread.getMessages().forEach(function(msg) {
      if (durduruldu) return;
      if (!msg.isUnread()) return;
      if (Date.now() - baslangic > SURE_SINIRI_MS) {
        durduruldu = true;
        Logger.log("⏱ Süre sınırına yaklaşıldı, kalan mailler bir sonraki çalıştırmaya bırakıldı.");
        return;
      }
      try {
        var sonuc = isleMail(msg, shFiy, shLog);
        // ★ DÜZELTME: önceden isleMail içeride LINK_YOK/SAYFA_HATASI/PARSE_BASARISIZ
        // gibi durumlarda hata fırlatmadan sessizce dönüyordu; bu satır her zaman
        // çalıştığı için başarısız işlenen faturalar da "okundu" sayılıp bir daha
        // hiç taranmıyordu. Artık sadece gerçekten sonuçlanan (başarılı, zaten
        // işlenmiş, filtre dışı) mailler okundu işaretleniyor — teknik hatalarda
        // mail okunmadı kalıp bir sonraki çalıştırmada otomatik tekrar denenecek.
        if (sonuc && sonuc.markRead) msg.markRead();
      } catch(e) {
        Logger.log("HATA: " + e.message);
        logYaz(shLog, "?", msg.getFrom(), msg.getDate(), "HATA", e.message);
      }
    });
  });
}

function isleMail(msg, shFiy, shLog) {
  var konu  = msg.getSubject();
  var gond  = msg.getFrom();
  var tarih = msg.getDate();
  var body  = msg.getBody();

  Logger.log("─── " + konu);

  var fatNo = "";
  var m = konu.match(/- ([A-Z]{2,6}\d{8,16}) -/);
  if (m) fatNo = m[1];

  if (FATURA_FILTRE.length > 0 && fatNo) {
    var gecerli = FATURA_FILTRE.some(function(on) {
      return fatNo.toUpperCase().indexOf(on.toUpperCase()) === 0;
    });
    if (!gecerli) { Logger.log("Filtre dışı: " + fatNo); return { markRead: true }; }
  }

  if (fatNo && faturaIslendiMi(shLog, fatNo)) {
    Logger.log("Zaten işlendi: " + fatNo);
    return { markRead: true };
  }

  var edmLink = linkBulMailden(body);
  if (!edmLink) {
    logYaz(shLog, fatNo, gond, tarih, "LINK_YOK", "Mail'de link bulunamadı");
    return { markRead: false };
  }

  var anaHTML = sayfaCek(edmLink);
  if (!anaHTML) {
    logYaz(shLog, fatNo, gond, tarih, "SAYFA_HATASI", "Ana sayfa açılamadı");
    return { markRead: false };
  }

  var htmlLink = htmlLinkBul(anaHTML);
  if (!htmlLink) {
    logYaz(shLog, fatNo, gond, tarih, "HTML_LINK_YOK", "Fatura HTML linki bulunamadı");
    return { markRead: false };
  }

  var faturaHTML = sayfaCek(htmlLink);
  if (!faturaHTML) {
    logYaz(shLog, fatNo, gond, tarih, "FATURA_HTML_HATASI", "Fatura HTML açılamadı");
    return { markRead: false };
  }

  var driveLink = "";
  try {
    driveLink = faturaHtmliPDFKaydet(faturaHTML, fatNo, tarih);
    Logger.log("PDF kaydedildi: " + driveLink);
  } catch(pdfErr) {
    Logger.log("PDF kayıt hatası: " + pdfErr.message);
  }

  // ★ DÜZELTME: htmlLink (geçici/kısa ömürlü) yerine edmLink (kalıcı) kaydediliyor.
  var urunler = faturaParseEt(faturaHTML, fatNo, gond, tarih, driveLink, edmLink);
  if (!urunler || urunler.length === 0) {
    // ★ GEÇİCİ TEŞHİS: PARSE_BASARISIZ nedenini anlamak için ham HTML hakkında
    // kısa bir özet DETAY'a ekleniyor (uzunluk, null byte var mı, <tr> sayısı,
    // ilk 60 karakter). Kök neden bulunduktan sonra bu satır kaldırılabilir.
    var teshis = "len=" + faturaHTML.length +
      " null=" + (/\x00/.test(faturaHTML)) +
      " tr=" + ((faturaHTML.match(/<tr[\s>]/gi) || []).length) +
      " ilk60=" + faturaHTML.substring(0, 60).replace(/[\r\n\t]/g, " ");
    logYaz(shLog, fatNo, gond, tarih, "PARSE_BASARISIZ", "Ürün parse edilemedi | " + teshis);
    return { markRead: false };
  }

  // ★ DÜZELTME: nakliyeDagit() (palet/nakliye maliyetini m²'ye göre dağıtma) SADECE
  // Canyap faturaları için tasarlanmış özel bir kriterdir — FATURA_FILTRE boşaltılıp
  // tüm tedarikçilerin faturaları kabul edilmeye başlandığında bu adım da yanlışlıkla
  // TÜM tedarikçilere uygulanır hale gelmişti (örn. bir tedarikçinin ürün adında
  // "TAŞIMA" veya "SEVK" kelimesi geçerse o kalem nakliye sanılıp seramiklere
  // dağıtılabiliyordu). Artık sadece tedarikçi CANYAP ise çalışıyor, diğer tüm
  // tedarikçilerin faturaları standart (netFiyat = maliyetFiyat, nakliyePayi = 0)
  // şekilde kalıyor.
  var canyapFaturasi = urunler.length > 0 && urunler[0].tedarikci === "CANYAP";
  urunler = canyapFaturasi ? nakliyeDagit(urunler) : urunler;
  urunler.forEach(function(u) { fiyatYaz(shFiy, u); });
  logYaz(shLog, fatNo, gond, tarih, "BASARILI", urunler.length + " ürün");
  Logger.log("✅ " + fatNo + " → " + urunler.length + " ürün");
  return { markRead: true };
}

function faturaHtmliPDFKaydet(html, fatNo, tarih) {
  var klasorler = DriveApp.getFoldersByName(PDF_FOLDER);
  var klasor = klasorler.hasNext() ? klasorler.next() : DriveApp.createFolder(PDF_FOLDER);

  var d = tarih instanceof Date ? tarih : new Date(tarih);
  var ayKlasorAdi = Utilities.formatDate(d, "Europe/Istanbul", "yyyy-MM");
  var ayKlasorler = klasor.getFoldersByName(ayKlasorAdi);
  var ayKlasoru = ayKlasorler.hasNext() ? ayKlasorler.next() : klasor.createFolder(ayKlasorAdi);

  var mevcutlar = ayKlasoru.getFilesByName(fatNo + ".pdf");
  if (mevcutlar.hasNext()) return mevcutlar.next().getUrl();

  var htmlBlob = Utilities.newBlob(html, "text/html", fatNo + ".html");
  var token = ScriptApp.getOAuthToken();

  var meta = JSON.stringify({
    name: fatNo,
    mimeType: "application/vnd.google-apps.document",
    parents: [ayKlasoru.getId()]
  });

  var boundary = "fatura_pdf_boundary";
  var part1 = "--" + boundary + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + meta + "\r\n";
  part1 += "--" + boundary + "\r\nContent-Type: text/html\r\n\r\n";
  var part2 = "\r\n--" + boundary + "--";

  var allBytes = Utilities.newBlob(part1).getBytes()
    .concat(htmlBlob.getBytes())
    .concat(Utilities.newBlob(part2).getBytes());

  var uploadResp = UrlFetchApp.fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
    {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + token,
        "Content-Type": "multipart/related; boundary=" + boundary
      },
      payload: allBytes,
      muteHttpExceptions: true
    }
  );

  var uploadResult = JSON.parse(uploadResp.getContentText());
  if (!uploadResult.id) throw new Error("Docs yükleme hatası: " + uploadResp.getContentText().substring(0,200));

  var docId = uploadResult.id;
  var pdfResp = UrlFetchApp.fetch(
    "https://www.googleapis.com/drive/v3/files/" + docId + "/export?mimeType=application/pdf",
    { method: "GET", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true }
  );

  var pdfBlob = pdfResp.getBlob().setName(fatNo + ".pdf");
  var pdfFile = ayKlasoru.createFile(pdfBlob);
  try { DriveApp.getFileById(docId).setTrashed(true); } catch(e) {}
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // /preview formatı direkt tarayıcıda PDF açar, izin sorunu olmaz
  return "https://drive.google.com/file/d/" + pdfFile.getId() + "/preview";
}

function nakliyeDagit(urunler) {
  var nakliyeKalemleri = [];
  var normalUrunler = [];

  urunler.forEach(function(u) {
    var adUpper = (u.stokAdi || "").toUpperCase();
    var isNakliyeKw  = NAKLIYE_KW.some(function(kw) { return adUpper.indexOf(kw) > -1; });
    var isNakliyeKod = NAKLIYE_KODLAR.indexOf(u.stokKodu) > -1;
    if (isNakliyeKw || isNakliyeKod) {
      nakliyeKalemleri.push(u);
      Logger.log("Nakliye/Palet tespit: " + u.stokAdi + " → " + u.netFiyat + " TL");
    } else {
      normalUrunler.push(u);
    }
  });

  if (nakliyeKalemleri.length === 0) return urunler;

  var toplamNakliye = 0;
  nakliyeKalemleri.forEach(function(u) {
    toplamNakliye += (u.netFiyat || u.birimFiyat) * (u.miktar || 1);
  });
  Logger.log("Toplam nakliye+palet: " + toplamNakliye + " TL");

  var seramikler = normalUrunler.filter(function(u) {
    var adUpper = (u.stokAdi || "").toUpperCase();
    return isSeramik(adUpper) && u.miktar > 0;
  });

  if (seramikler.length === 0) {
    Logger.log("Seramik bulunamadı, nakliye dağıtılamadı");
    return urunler;
  }

  var toplamM2 = 0;
  seramikler.forEach(function(u) { toplamM2 += u.miktar; });
  Logger.log("Toplam m²: " + toplamM2);
  if (toplamM2 === 0) return urunler;

  var nakliyeM2 = toplamNakliye / toplamM2;
  Logger.log("Nakliye/m²: " + nakliyeM2.toFixed(4) + " TL");

  normalUrunler.forEach(function(u) {
    var adUpper = (u.stokAdi || "").toUpperCase();
    if (isSeramik(adUpper) && u.miktar > 0) {
      u.nakliyePayi  = Math.round(nakliyeM2 * 10000) / 10000;
      u.maliyetFiyat = Math.round((u.netFiyat + u.nakliyePayi) * 10000) / 10000;
      Logger.log(u.stokKodu + " → nakliye:" + u.nakliyePayi + " maliyet:" + u.maliyetFiyat);
    }
  });

  return normalUrunler;
}

function isSeramik(adUpper) {
  return SERAMIK_KW.some(function(kw) { return adUpper.indexOf(kw) > -1; });
}

// ★ YENİ: Fatura HTML'inden doğru tarihi çıkarır — "Vade Tarihi"/"Ödeme Tarihi"ni asla kabul etmez.
// Hem faturaParseEt() hem de tarihleriYenidenCek() onarım fonksiyonu tarafından kullanılır.
function tarihCikarHTMLden(html, yedekTarih) {
  // ★ DÜZELTME: Faturada "Fatura Tarihi" TİRE (-) ile ("03-07-2026"), ama
  // "Vade Tarihi" / "Ödeme Tarihi" NOKTA (.) ile ("31.10.2026") yazılıyormuş.
  // Regex sadece nokta arıyordu, bu yüzden doğru (tireli) tarihi hiç
  // görmüyor, hep yanlış (noktalı) vade/ödeme tarihini buluyordu.
  // Artık her iki ayırıcı da destekleniyor: [.\-]
  var etiketDatePat = /([A-Za-zÇĞİÖŞÜçğıöşü ]{0,40})Tarihi[\s\S]{0,250}?(\d{2})[.\-](\d{2})[.\-](\d{4})/g;
  var m, adaylar = [];
  while ((m = etiketDatePat.exec(html)) !== null) {
    adaylar.push({
      etiket: m[1].trim().toUpperCase(),
      tarih: m[2] + "/" + m[3] + "/" + m[4]
    });
  }
  function harici(a) {
    return a.etiket.indexOf("VADE") === -1 &&
      a.etiket.indexOf("ÖDEME") === -1 &&
      a.etiket.indexOf("ODEME") === -1;
  }
  var secilen = adaylar.find(function(a) {
    return harici(a) && (
      a.etiket.indexOf("FATURA") > -1 ||
      a.etiket.indexOf("DÜZENLEME") > -1 || a.etiket.indexOf("DUZENLEME") > -1 ||
      a.etiket.indexOf("SİPARİŞ") > -1 || a.etiket.indexOf("SIPARIS") > -1
    );
  });
  if (!secilen) {
    secilen = adaylar.find(harici);
  }
  if (secilen) return secilen.tarih;

  // Etiketli hiçbir tarih bulunamadıysa (beklenmedik sayfa yapısı): sayfadaki
  // ilk tarihi son çare olarak kullanılır — "bugün"e düşmeden önce.
  var genelM = html.match(/(\d{2})[.\-](\d{2})[.\-](\d{4})/);
  if (genelM) return genelM[1] + "/" + genelM[2] + "/" + genelM[3];

  var d = (yedekTarih instanceof Date) ? yedekTarih : new Date(yedekTarih);
  return Utilities.formatDate(d, "Europe/Istanbul", "dd/MM/yyyy");
}

function faturaParseEt(html, fatNo, gond, tarih, driveLink, edmLink) {
  var urunler = [];
  var now = new Date();

  var tedarikci = gond.replace(/"([^"]+)"[\s\S]*/, "$1").replace(/<[^>]+>/g,"").trim();
  if (fatNo && fatNo.toUpperCase().indexOf("CNY") === 0) tedarikci = "CANYAP";

  var fatTarih = tarihCikarHTMLden(html, tarih);

  var temizHTML = html;
  while (temizHTML.indexOf("<script") > -1) {
    var s = temizHTML.toLowerCase().indexOf("<script");
    var e = temizHTML.toLowerCase().indexOf("</script>", s);
    if (e === -1) break;
    temizHTML = temizHTML.substring(0, s) + temizHTML.substring(e + 9);
  }
  while (temizHTML.indexOf("<style") > -1) {
    var s2 = temizHTML.toLowerCase().indexOf("<style");
    var e2 = temizHTML.toLowerCase().indexOf("</style>", s2);
    if (e2 === -1) break;
    temizHTML = temizHTML.substring(0, s2) + temizHTML.substring(e2 + 8);
  }

  function sayiyaCevir(str) {
    if (!str) return 0;
    // ★ DÜZELTME: eskiden TÜM string'deki rakam/virgül/nokta karakterleri (aralarında
    // boşluk veya parantez olsa bile) tek bir sayıya birleştiriliyordu. Örn. hücre
    // içeriği "38,88 (2)" (miktarın yanında koli/paket sayısı parantez içinde) ise
    // eski mantık boşluk ve parantezi atıp "38,882" gibi YANLIŞ, birleşik bir sayı
    // üretiyordu (CNY2026000002227 faturasında miktarın 38,88 yerine 38,882 görünmesinin
    // kök nedeni buydu). Artık string içindeki İLK bitişik sayısal belirteç
    // yakalanıyor — boşluk/parantez ile ayrılmış başka bir sayı asıl sayıya karışmıyor.
    var m = String(str).match(/\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+(?:,\d+)?/);
    if (!m) return 0;
    var t = m[0];
    t = t.replace(/\.(?=\d{3}(?:[,]|$))/g, "");
    t = t.replace(",", ".");
    return parseFloat(t) || 0;
  }

  // Bir hücrenin (birim adı/ürün adı gibi metin değil) rakamsal bir değer olup
  // olmadığını anlamak için kullanılıyor — "TL" ve "%" gibi ekleri yok sayar.
  function hucreSayiMi(str) {
    if (!str) return false;
    var t = str.trim().replace(/\bTL\b/gi, "").replace(/%/g, "").trim();
    if (!/\d/.test(t)) return false;
    return !/[a-zA-ZÇĞİÖŞÜçğıöşü]/.test(t);
  }

  var trPat = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  var trM;
  var satirSayisi = 0;

  while ((trM = trPat.exec(temizHTML)) !== null) {
    var satir = trM[1];
    var tdler = [];
    var tdPat = /<td[^>]*>([\s\S]*?)<\/td>/gi;
    var tdM;
    while ((tdM = tdPat.exec(satir)) !== null) {
      var ic = tdM[1]
        .replace(/<[^>]+>/g, " ")
        .replace(/&nbsp;/g, " ")
        .replace(/&amp;/g, "&")
        .replace(/\s+/g, " ")
        .trim();
      // ★ DÜZELTME: boş hücreler ("İskonto Oranı" gibi bazı sütunlar iskonto
      // yokken boş gelir) artık diziden SİLİNMİYOR — eskiden silinmesi, sütun
      // pozisyonuna dayalı ayrıştırmalarda (siraNoStil vb.) tüm sütunların
      // kaymasına ve hatalı PARSE_BASARISIZ sonucuna yol açıyordu.
      tdler.push(ic);
    }
    if (tdler.length < 3) continue;
    satirSayisi++;

    var stokKoduIdx = -1;
    var stokKodu = "", stokAdi = "";
    var birlesikStil = false;

    // 1) Canyap stili (değiştirilmedi): kod tek başına bir hücrede (8-15 haneli sadece rakam)
    for (var i = 0; i < tdler.length; i++) {
      if (/^\d{8,15}$/.test(tdler[i].trim())) {
        stokKodu = tdler[i].trim();
        stokKoduIdx = i;
        break;
      }
    }

    // 2) ★ YENİ: kod ve ürün adı aynı hücrede birleşik gelen şablon
    //    (örn. "310200500022 - KALE VERSUS İNCE GÖM REZ HELA TAŞI 79099")
    //    — Canyap dışı bazı tedarikçilerde (EDM Bilişim üzerinden farklı firma) görüldü.
    if (!stokKodu) {
      for (var j = 0; j < tdler.length; j++) {
        var bm = tdler[j].trim().match(/^(\d{8,15})\s*[-–]\s*(.+)$/);
        if (bm) {
          stokKodu = bm[1];
          stokAdi = bm[2].trim();
          stokKoduIdx = j;
          birlesikStil = true;
          break;
        }
      }
    }

    // 3) ★ YENİ: STOK KODU hiç yok — "Sıra No | Mal Hizmet | Miktar Birim | Birim Fiyat |
    //    İskonto Oranı | İskonto Tutarı | KDV Oranı | KDV Tutarı | [Diğer Vergiler |]
    //    Mal Hizmet Tutarı" şablonu (örn. GPD, GEF, MUS, PAA... gibi tedarikçilerde
    //    görüldü — bazılarında "Diğer Vergiler" sütunu var, bazılarında yok, bu yüzden
    //    Mal Hizmet Tutarı sabit index yerine HER ZAMAN son sütundan okunuyor).
    //    Ayırt edici: ilk hücre kısa bir sıra numarası (1-4 haneli), ikinci hücre metin
    //    (ürün adı) ve satırda en az 9 hücre var — bu kombinasyon faturanın başka
    //    tablolarında (adres/toplam vb.) pratikte oluşmuyor.
    var siraNoStil = false;
    if (!stokKodu && tdler.length >= 9 &&
        /^\d{1,4}$/.test(tdler[0].trim()) && !hucreSayiMi(tdler[1])) {
      siraNoStil = true;
      stokKodu = ""; // bu şablonda stok kodu gerçekten yok
      stokAdi  = tdler[1].trim();
      stokKoduIdx = 1;
    }

    // 4) ★ YENİ: kısa/nokta formatlı kod, ürün adından AYRI hücrede
    //    (örn. "02.005", "08.001" — FME/FMF Parke Mobilya gibi tedarikçilerde görüldü).
    //    "SIRA | KODU | AÇIKLAMASI | MIKTAR | FIYAT | KDV% | KDV | TUTAR" — iskonto sütunu
    //    yok, TUTAR = miktar×fiyat (KDV hariç), her zaman son sütun.
    //    Ayırt edici: 1. hücre sıra no, 2. hücre "12.345" gibi nokta içeren kısa kod
    //    (hucreSayiMi ile sayısal görünür ama 8-15 haneli DEĞİL), 3. hücre ürün adı (metin).
    var kisaKoduStil = false;
    if (!stokKodu && !siraNoStil && tdler.length >= 8 &&
        /^\d{1,4}$/.test(tdler[0].trim()) &&
        /^\d{1,4}\.\d{1,5}$/.test(tdler[1].trim()) &&
        !hucreSayiMi(tdler[2])) {
      kisaKoduStil = true;
      stokKodu = tdler[1].trim();
      stokAdi  = tdler[2].trim();
      stokKoduIdx = 2;
    }

    if (!stokKodu && !siraNoStil && !kisaKoduStil) continue;

    var miktar, birimFiyat, iskonto, netFiyat, kdv;

    if (kisaKoduStil) {
      // "13,76 M²" gibi miktar+birim birleşik hücreyi ayır — baştaki sayıyı al.
      var mbM3 = tdler[3].trim().match(/^([\d.,]+)/);
      miktar = mbM3 ? sayiyaCevir(mbM3[1]) : 0;
      birimFiyat = sayiyaCevir(tdler[4] || "");
      iskonto = 0; // bu şablonda iskonto sütunu hiç yok
      kdv = sayiyaCevir(tdler[5] || "");
      var tutarNet = sayiyaCevir(tdler[tdler.length - 1] || ""); // KDV hariç, miktar×fiyat
      netFiyat = miktar > 0 ? Math.round((tutarNet / miktar) * 10000) / 10000 : birimFiyat;
    } else if (siraNoStil) {
      // "50 Adet" gibi miktar+birim birleşik hücreyi ayır — baştaki sayıyı al.
      var mbM = tdler[2].trim().match(/^([\d.,]+)/);
      miktar = mbM ? sayiyaCevir(mbM[1]) : 0;
      birimFiyat = sayiyaCevir(tdler[3] || "");
      iskonto    = sayiyaCevir(tdler[4] || ""); // "%54,00" — sayiyaCevir % işaretini zaten süzüyor
      kdv        = sayiyaCevir(tdler[6] || "");
      var malHizmetTutari = sayiyaCevir(tdler[tdler.length - 1] || ""); // her zaman son sütun
      netFiyat = miktar > 0 ? Math.round((malHizmetTutari / miktar) * 10000) / 10000 : birimFiyat * (1 - iskonto / 100);
    } else if (birlesikStil) {
      // Kod hücresinden sonraki hücrelerden sayısal olanları sırayla topla
      // (aradaki "Adet" gibi birim hücresi otomatik atlanır).
      var sayilar = [];
      for (var k = stokKoduIdx + 1; k < tdler.length; k++) {
        if (hucreSayiMi(tdler[k])) sayilar.push(sayiyaCevir(tdler[k]));
      }
      if (sayilar.length >= 6) {
        // Görülen sıra: MİKTAR, FİYAT, İSKONTO ORANI, İSKONTO TUTARI, KDV TUTARI, TUTAR
        miktar = sayilar[0];
        birimFiyat = sayilar[1];
        iskonto = sayilar[2];
        var kdvTutari   = sayilar[4];
        var satirTutari = sayilar[5];
        netFiyat = miktar > 0 ? Math.round((satirTutari / miktar) * 10000) / 10000 : birimFiyat * (1 - iskonto / 100);
        // KDV_ORANI alanı önceden bazı şablonlarda yanlışlıkla tutar (₺) olarak
        // kaydediliyordu; burada gerçek oran (%) hesaplanıyor.
        kdv = satirTutari > 0 ? Math.round((kdvTutari / satirTutari) * 10000) / 100 : 0;
      } else {
        continue; // beklenmeyen kolon sayısı — yanlış veri yazmamak için satır atlanıyor
      }
    } else {
      // Eski (Canyap) mantığı — hiç değiştirilmedi
      stokAdi = stokKoduIdx + 1 < tdler.length ? tdler[stokKoduIdx + 1] : "";
      miktar     = sayiyaCevir(tdler[stokKoduIdx + 2] || "");
      birimFiyat = sayiyaCevir(tdler[stokKoduIdx + 3] || "");
      iskonto    = sayiyaCevir(tdler[stokKoduIdx + 4] || "");

      if (iskonto > 100) {
        Logger.log("Kolon kayması: " + tdler.join(" | "));
        birimFiyat = iskonto;
        iskonto    = sayiyaCevir(tdler[stokKoduIdx + 5] || "");
        netFiyat   = sayiyaCevir(tdler[stokKoduIdx + 6] || "");
        kdv        = sayiyaCevir(tdler[stokKoduIdx + 7] || "");
      } else {
        netFiyat   = sayiyaCevir(tdler[stokKoduIdx + 5] || "");
        kdv        = sayiyaCevir(tdler[stokKoduIdx + 6] || "");
      }

      if (netFiyat === 0 && birimFiyat > 0) {
        netFiyat = birimFiyat * (1 - iskonto / 100);
        netFiyat = Math.round(netFiyat * 10000) / 10000;
      }
    }

    if (birimFiyat > 0 || netFiyat > 0) {
      urunler.push({
        stokKodu:     stokKodu,
        stokAdi:      stokAdi,
        miktar:       miktar,
        birimFiyat:   birimFiyat,
        iskonto:      iskonto,
        netFiyat:     netFiyat,
        nakliyePayi:  0,
        maliyetFiyat: netFiyat,
        kdvOrani:     kdv,
        faturaNo:     fatNo,
        fatTarih:     fatTarih,
        tedarikci:    tedarikci,
        driveLink:    driveLink || "",
        edmLink:      edmLink  || "",
        islemZamani:  now
      });
    }
  }

  Logger.log("Parse: " + satirSayisi + " satır, " + urunler.length + " ürün");
  return urunler;
}

// ★ YENİ: FATURAFIYAT sayfasına MIKTAR kolonunu ekler (bir kereye mahsus migrasyon,
// tekrar çalıştırmak zararsız — kolon zaten varsa hiçbir şey yapmaz).
function miktarKolonuEkle(sheet) {
  var baslik = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (baslik.indexOf("MIKTAR") > -1) return; // zaten eklenmiş
  sheet.insertColumnAfter(2); // STOK_ADI'nın (2. kolon) hemen sağına ekle
  sheet.getRange(1, 3).setValue("MIKTAR").setFontWeight("bold").setBackground("#e8edf5");
  Logger.log("✅ MIKTAR kolonu eklendi (3. kolon).");
}

function fiyatYaz(sheet, u) {
  var fatTarihStr = u.fatTarih;
  if (fatTarihStr && fatTarihStr.toString().indexOf("GMT") > -1) {
    var d = new Date(fatTarihStr);
    fatTarihStr = Utilities.formatDate(d, "Europe/Istanbul", "dd/MM/yyyy");
  }

  var satir = [
    u.stokKodu, u.stokAdi, u.miktar || 0, u.birimFiyat, u.iskonto, u.netFiyat,
    u.nakliyePayi || 0, u.maliyetFiyat || u.netFiyat, u.kdvOrani,
    u.faturaNo, fatTarihStr, u.tedarikci,
    u.driveLink || "", u.edmLink || "",
    Utilities.formatDate(u.islemZamani, "Europe/Istanbul", "dd/MM/yyyy HH:mm")
  ];

  sheet.insertRowAfter(1);

  // FATURA_TARIHI kolonunu dinamik bul (MIKTAR eklenince kolon numarası kaydı) ve
  // düz metin formatına zorla — aksi halde Sheets "dd/MM/yyyy" metnini otomatik
  // olarak gerçek bir tarih hücresine çevirip bozuyordu.
  var baslik = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var iFTar = baslik.indexOf("FATURA_TARIHI");
  if (iFTar > -1) sheet.getRange(2, iFTar + 1).setNumberFormat("@");

  sheet.getRange(2, 1, 1, satir.length).setValues([satir]);
}

// ═══════════════════════════════════════════════════════════════════
//  2026 FATURALARINI ÇEK (okunmuş/okunmamış fark etmez)
// ═══════════════════════════════════════════════════════════════════

function efatura2026Getir() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var shFiy = getOrCreateSheet(ss, SHEET_FIYAT, [
    "STOK_KODU","STOK_ADI","MIKTAR","BIRIM_FIYAT","ISKONTO","NET_FIYAT",
    "NAKLIYE_PAYI","MALIYET_FIYAT","KDV_ORANI",
    "FATURA_NO","FATURA_TARIHI","TEDARIKCI",
    "FATURA_LINK","EDM_LINK","ISLEM_ZAMANI"
  ]);
  miktarKolonuEkle(shFiy);
  var shLog = getOrCreateSheet(ss, SHEET_LOG,
    ["FATURA_NO","GONDEREN","TARIH","DURUM","DETAY","ISLEM_ZAMANI"]);

  var query2026 = 'subject:"e-Faturanız var"';
  var threads = GmailApp.search(query2026, 0, 500);

  if (threads.length === 0) {
    Logger.log("Hiç e-fatura bulunamadı.");
    return;
  }
  Logger.log(threads.length + " e-fatura thread bulundu. CNY2026 filtreleniyor...");

  var islenen = 0;
  var atlanan = 0;

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      var konu = msg.getSubject();
      if (konu.indexOf("CNY2026") === -1) return;

      try {
        var fatNo = "";
        var m = konu.match(/- ([A-Z]{2,6}\d{8,16}) -/);
        if (m) fatNo = m[1];

        if (fatNo && faturaIslendiMi(shLog, fatNo)) {
          Logger.log("Zaten işlenmiş, atlanıyor: " + fatNo);
          atlanan++;
          return;
        }

        isleMail(msg, shFiy, shLog);
        islenen++;
        try { msg.markRead(); } catch(e) {}
      } catch(e) {
        Logger.log("HATA (2026): " + e.message);
        logYaz(shLog, "?", msg.getFrom(), msg.getDate(), "HATA", e.message);
      }
    });
  });

  Logger.log("✅ 2026 tamamlandı: " + islenen + " yeni fatura işlendi, " + atlanan + " önceden işlenmiş atlandı.");
}

function manuelEfatura2026() {
  try {
    efatura2026Getir();
    return { ok: true, mesaj: "2026 e-fatura çekimi tamamlandı" };
  } catch(e) {
    return { ok: false, hata: e.message };
  }
}

function manuelEfatura() {
  try {
    efaturaOku();
    return { ok: true, mesaj: "E-fatura çekimi tamamlandı" };
  } catch(e) {
    return { ok: false, hata: e.message };
  }
}

function test2026() {
  Logger.log("=== 2026 FATURA MANUEL TEST ===");
  efatura2026Getir();
  Logger.log("=== BİTTİ ===");
}

// ═══════════════════════════════════════════════════════════════════
//  YARDIMCILAR (FATURA)
// ═══════════════════════════════════════════════════════════════════

function sayfaCek(url) {
  try {
    var resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" }
    });
    if (resp.getResponseCode() !== 200) {
      Logger.log("HTTP " + resp.getResponseCode() + " — " + url);
      return null;
    }
    return decodeSayfaIcerigi(resp);
  } catch(e) {
    Logger.log("Fetch hatası: " + e.message);
    return null;
  }
}

// ★ DÜZELTME: EDM Bilişim'in bazı fatura sayfaları (<meta charset="utf-16">
// ile işaretli — TOS/GL5/GEF/EK02/EFS/PAA/CNA/DPT gibi Canyap dışı tedarikçi
// şablonları) gerçekten UTF-16 bayt dizisi olarak geliyor. Eskiden içerik
// sabit "UTF-8" ile decode ediliyordu; bu da <tr>/<td> etiketlerini bozup
// faturaParseEt'in hiç satır bulamamasına (PARSE_BASARISIZ) yol açıyordu.
// Artık ham baytlardaki BOM'a (yoksa meta etiketine) bakılıp doğru encoding
// ile decode ediliyor.
function decodeSayfaIcerigi(resp) {
  var bytes = resp.getContent();
  if (bytes && bytes.length >= 2) {
    var b0 = bytes[0] & 0xFF, b1 = bytes[1] & 0xFF;
    if (b0 === 0xFF && b1 === 0xFE) return resp.getContentText("UTF-16LE");
    if (b0 === 0xFE && b1 === 0xFF) return resp.getContentText("UTF-16BE");
  }
  var utf8 = resp.getContentText("UTF-8");
  // BOM yoksa ama meta etiketi utf-16 diyorsa (BOM'suz UTF-16 ihtimali) yine dene.
  // UTF-16 baytları UTF-8 olarak okununca harfler arasına \x00 sıkışır, bu yüzden
  // kontrol öncesi \x00'ları temizliyoruz.
  if (/charset\s*=\s*["']?utf-16/i.test(utf8.replace(/\x00/g, ""))) {
    try {
      var alt = resp.getContentText("UTF-16LE");
      if (/<html/i.test(alt)) return alt;
    } catch(e) {}
  }
  return utf8;
}

function linkBulMailden(body) {
  var patterns = [
    /https:\/\/view\.edmbilisim\.com\.tr\/fatura\/ViewInvoice\/[^\s"'<>]+/gi,
    /href="(https?:\/\/[^"]*edmbilisim[^"]*ViewInvoice[^"]*)"/i,
  ];
  for (var i = 0; i < patterns.length; i++) {
    var found = body.match(patterns[i]);
    if (found) return found[1] || found[0];
  }
  return null;
}

function htmlLinkBul(html) {
  var m = html.match(/href="(\/fatura\/tmp-ef\/[^"]+\.html)"/i) ||
          html.match(/(\/fatura\/tmp-ef\/[^\s"'<>]+\.html)/i);
  if (m) return EDM_BASE + m[1];
  return null;
}

function faturaIslendiMi(shLog, fatNo) {
  try {
    if (!shLog) return false;
    var logData = shLog.getDataRange().getValues();
    if (logData.length <= 1) return false;

    var logBasarili = false;
    for (var i = 1; i < logData.length; i++) {
      if (String(logData[i][0]) === String(fatNo) && String(logData[i][3]) === "BASARILI") {
        logBasarili = true;
        break;
      }
    }
    if (!logBasarili) return false;

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var shFiy = ss.getSheetByName(SHEET_FIYAT);
    if (!shFiy) return false;
    var fiyData = shFiy.getDataRange().getValues();
    if (fiyData.length < 2) return false;

    var faturaNoIdx = fiyData[0].indexOf("FATURA_NO");
    if (faturaNoIdx === -1) return false;

    for (var j = 1; j < fiyData.length; j++) {
      if (String(fiyData[j][faturaNoIdx]) === String(fatNo)) {
        return true;
      }
    }

    Logger.log("Logda BASARILI ama fiyatta veri yok, tekrar işlenecek: " + fatNo);
    return false;

  } catch(e) {
    Logger.log("faturaIslendiMi hatası: " + e.message);
    return false;
  }
}

function logYaz(sh, fatNo, gond, tarih, durum, detay) {
  sh.appendRow([
    fatNo, gond,
    Utilities.formatDate(new Date(tarih), "Europe/Istanbul", "dd/MM/yyyy HH:mm"),
    durum, detay,
    Utilities.formatDate(new Date(), "Europe/Istanbul", "dd/MM/yyyy HH:mm")
  ]);
}

// ═══════════════════════════════════════════════════════════════════
//  STOK YÖNETİMİ
// ═══════════════════════════════════════════════════════════════════

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e8edf5");
  }
  return sheet;
}

function findColIndex(headers, candidates) {
  for (const c of candidates) {
    const idx = headers.indexOf(c);
    if (idx >= 0) return idx;
  }
  return -1;
}

function parseDateFromFilename(name) {
  let m = name.match(/(\d{2})[._\-\s](\d{2})[._\-\s](\d{4})/);
  if (m) return m[3] + "-" + m[2].padStart(2,"0") + "-" + m[1].padStart(2,"0");
  m = name.match(/(\d{4})[._\-\s](\d{2})[._\-\s](\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];
  m = name.match(/(\d{2})[._\-\s](\d{2})/);
  if (m) return new Date().getFullYear() + "-" + m[2].padStart(2,"0") + "-" + m[1].padStart(2,"0");
  return new Date().toISOString().split("T")[0];
}

function stokLogEntry(tip, detay, tarih) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, STOK_SHEETS.log, ["TARİH", "TİP", "DETAY"]);
  sheet.appendRow([tarih || new Date(), tip, detay]);
}

function stokLogError(err) {
  try { stokLogEntry("HATA", err.message, new Date()); } catch(e) {}
}

function getFiyatlar() {
  const fiyatMap    = {};
  const gecmisMap   = {};

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_FIYAT);
    if (!sh) return { fiyatMap, gecmisMap };

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { fiyatMap, gecmisMap };

    const h = data[0];
    const colKod  = h.indexOf("STOK_KODU");
    const colFiy  = h.indexOf("BIRIM_FIYAT");
    const colIsk  = h.indexOf("ISKONTO");
    const colNet  = h.indexOf("NET_FIYAT");
    const colNav  = h.indexOf("NAKLIYE_PAYI");
    const colMal  = h.indexOf("MALIYET_FIYAT");
    const colKdv  = h.indexOf("KDV_ORANI");
    const colFNo  = h.indexOf("FATURA_NO");
    const colFTar = h.indexOf("FATURA_TARIHI");
    const colTed  = h.indexOf("TEDARIKCI");
    const colLnk  = h.indexOf("FATURA_LINK");
    const colEdm  = h.indexOf("EDM_LINK");

    data.slice(1).forEach(row => {
      const kod = String(row[colKod] || "").trim();
      if (!kod) return;

      // ★ Ekstra güvence: FATURA_TARIHI hücresi (eski bozuk kayıtlarda olabileceği gibi)
      // gerçek bir Date nesnesi olarak gelirse, String() ile çirkin toString() çıktısı
      // yerine düzgün "dd/MM/yyyy" formatına çeviriyoruz.
      const rawTarih = row[colFTar];
      const fatTarihStr = (rawTarih instanceof Date)
        ? Utilities.formatDate(rawTarih, "Europe/Istanbul", "dd/MM/yyyy")
        : String(rawTarih || "");

      const kayit = {
        birimFiyat:   parseFloat(row[colFiy]) || 0,
        iskonto:      parseFloat(row[colIsk]) || 0,
        netFiyat:     colNet >= 0 ? parseFloat(row[colNet]) || 0 : 0,
        nakliyePayi:  colNav >= 0 ? parseFloat(row[colNav]) || 0 : 0,
        maliyetFiyat: colMal >= 0 ? parseFloat(row[colMal]) || 0 : 0,
        kdvOrani:     parseFloat(row[colKdv]) || 0,
        faturaNo:     String(row[colFNo]  || ""),
        fatTarih:     fatTarihStr,
        tedarikci:    String(row[colTed]  || ""),
        faturaLink:   colLnk >= 0 ? String(row[colLnk] || "") : "",
        edmLink:      colEdm >= 0 ? String(row[colEdm] || "") : "",
      };

      if (!fiyatMap[kod]) fiyatMap[kod] = kayit;
      if (!gecmisMap[kod]) gecmisMap[kod] = [];
      gecmisMap[kod].push(kayit);
    });

    Logger.log("Fiyat verisi okundu: " + Object.keys(fiyatMap).length + " ürün, " +
      Object.values(gecmisMap).reduce((t, a) => t + a.length, 0) + " toplam kayıt");

  } catch(e) {
    Logger.log("Fiyat okuma hatası: " + e.message);
  }

  return { fiyatMap, gecmisMap };
}

function parseExcelFile(excelFile) {
  const token    = ScriptApp.getOAuthToken();
  const boundary = "stok_boundary_xyz";

  const metaJson = JSON.stringify({
    name: "temp_stok_" + Date.now(),
    mimeType: "application/vnd.google-apps.spreadsheet"
  });

  const part1 = "--" + boundary + "\r\n" +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    metaJson + "\r\n" +
    "--" + boundary + "\r\n" +
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n";

  const part3 = "\r\n--" + boundary + "--";

  const allBytes = Utilities.newBlob(part1).getBytes()
    .concat(excelFile.getBlob().getBytes())
    .concat(Utilities.newBlob(part3).getBytes());

  const response = UrlFetchApp.fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
    {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + token,
        "Content-Type": "multipart/related; boundary=" + boundary
      },
      payload: allBytes,
      muteHttpExceptions: true
    }
  );

  const result = JSON.parse(response.getContentText());
  if (!result.id) {
    throw new Error("Excel→Sheets dönüşüm hatası: " + response.getContentText().substring(0, 200));
  }

  const tempSS = SpreadsheetApp.openById(result.id);
  const sheet  = tempSS.getSheets()[0];
  const data   = sheet.getDataRange().getValues();

  try { DriveApp.getFileById(result.id).setTrashed(true); } catch(e) {}

  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).toUpperCase().trim().replace(/ /g, "_"));

  const colMap = {
    STOK_KODU:   findColIndex(headers, ["STOK_KODU","STOK KODU","KOD","STOKKODU"]),
    STOK_ADI:    findColIndex(headers, ["STOK_ADI","STOK ADI","AD","STOKADI","AÇIKLAMA"]),
    MERKEZ_DEPO: findColIndex(headers, ["MERKEZ_DEPO","MERKEZ DEPO","DEPO","MERKEZ"]),
    AYRILMIS:    findColIndex(headers, ["AYRILMIS","AYRILAN","AYRILMIŞ","REZERVE"]),
    SATILABILIR: findColIndex(headers, ["SATILABILIR","SATILABİLİR","NET","KULLANILABILIR"]),
  };

  return data.slice(1).map(row => {
    const kod = colMap.STOK_KODU >= 0 ? String(row[colMap.STOK_KODU] || "").trim() : "";
    if (!kod) return null;
    return {
      STOK_KODU:   kod,
      STOK_ADI:    colMap.STOK_ADI    >= 0 ? String(row[colMap.STOK_ADI]    || "").trim() : "",
      MERKEZ_DEPO: colMap.MERKEZ_DEPO >= 0 ? parseFloat(row[colMap.MERKEZ_DEPO] || 0) || 0 : 0,
      AYRILMIS:    colMap.AYRILMIS    >= 0 ? parseFloat(row[colMap.AYRILMIS]    || 0) || 0 : 0,
      SATILABILIR: colMap.SATILABILIR >= 0 ? parseFloat(row[colMap.SATILABILIR] || 0) || 0 : 0,
    };
  }).filter(r => r && r.STOK_KODU);
}

function saveStoklar(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, STOK_SHEETS.stoklar,
    ["STOK_KODU","STOK_ADI","MERKEZ_DEPO","AYRILMIS","SATILABILIR","TARIH"]);

  const tarih    = body.tarih || new Date().toISOString().split("T")[0];
  const yeniList = body.stoklar || [];

  const yeniMap = {};
  yeniList.forEach(r => { yeniMap[String(r.STOK_KODU)] = r; });

  const mevcutData = sheet.getDataRange().getValues();
  const mevcutMap  = {};
  if (mevcutData.length > 1) {
    mevcutData.slice(1).forEach(row => {
      if (row[0]) mevcutMap[String(row[0])] = row;
    });
  }

  const tumKodlar = new Set([...Object.keys(yeniMap), ...Object.keys(mevcutMap)]);
  const satirlar  = [];

  tumKodlar.forEach(kod => {
    if (yeniMap[kod]) {
      const r = yeniMap[kod];
      satirlar.push([r.STOK_KODU, r.STOK_ADI, r.MERKEZ_DEPO, r.AYRILMIS, r.SATILABILIR, tarih]);
    } else if (mevcutMap[kod]) {
      const r = mevcutMap[kod];
      satirlar.push([r[0], r[1], 0, 0, 0, tarih + " (pasif)"]);
    }
  });

  sheet.clearContents();
  sheet.appendRow(["STOK_KODU","STOK_ADI","MERKEZ_DEPO","AYRILMIS","SATILABILIR","TARIH"]);
  if (satirlar.length > 0) {
    sheet.getRange(2, 1, satirlar.length, 6).setValues(satirlar);
  }

  const pasif = satirlar.filter(r => String(r[5]).includes("pasif")).length;
  return { ok: true, count: yeniList.length, pasif: pasif };
}

function checkStokMail() {
  const threads = GmailApp.search(MAIL_QUERY_STOK, 0, 50);

  if (threads.length === 0) {
    Logger.log("Yeni stok maili yok.");
    return;
  }

  let label = GmailApp.getUserLabelByName("stok-islendi");
  if (!label) label = GmailApp.createLabel("stok-islendi");

  const stokFolder = getOrCreateFolder("STOK");
  const now = new Date();

  threads.forEach(thread => {
    const threadLabels = thread.getLabels().map(l => l.getName());
    if (threadLabels.indexOf("stok-islendi") > -1) return;

    thread.getMessages().forEach(msg => {
      const to = msg.getTo() || "";
      if (to.toLowerCase().indexOf("fincanlaryapi@gmail.com") === -1) return;

      const daysOld = (now - msg.getDate()) / (1000 * 60 * 60 * 24);
      if (daysOld > 30) return;

      msg.getAttachments().forEach(att => {
        const name = att.getName();
        var konuUpper = (msg.getSubject() || "").toUpperCase();
        var isCanyapMail = konuUpper.indexOf("CANYAP") > -1 || name.toUpperCase().indexOf("CANYAP") > -1 || name.toUpperCase().indexOf("STOK") > -1;
        if (!name.match(/\.xlsx?$/i) || !isCanyapMail) return;

        Logger.log("Stok maili bulundu: " + name + " | Konu: " + msg.getSubject());

        const existing = stokFolder.getFilesByName(name);
        while (existing.hasNext()) existing.next().setTrashed(true);
        const savedFile = stokFolder.createFile(att);

        try {
          const stoklar = parseExcelFile(savedFile);
          if (stoklar.length > 0) {
            const tarih = parseDateFromFilename(name);
            saveStoklar({ stoklar: stoklar, tarih: tarih });
            stokLogEntry("otomatik", name + " → " + stoklar.length + " ürün", new Date());
            Logger.log("✅ Stok işlendi: " + name + " (" + stoklar.length + " ürün)");
          }
        } catch(err) {
          stokLogEntry("hata", name + ": " + err.message, new Date());
        }
      });
      thread.addLabel(label);
    });
  });
}

// ═══════════════════════════════════════════════════════════════════
//  WEB API
// ═══════════════════════════════════════════════════════════════════

function doGet(e) {
  if (e.parameter && e.parameter.action) {
    return handleRequest(e);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('CANYAP Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const body   = e.postData ? JSON.parse(e.postData.contents) : e.parameter;
    const action = body.action;
    let result;

    switch (action) {
      case "getAll":         result = getAllData();           break;
      case "saveNot":        result = saveNot(body);          break;
      case "saveEtiket":     result = saveEtiket(body);       break;
      case "saveUrunEt":     result = saveUrunEt(body);       break;
      case "saveMarka":      result = saveMarka(body);        break;
      case "saveAyar":       result = saveAyar(body);         break;
      case "saveStoklar":    result = saveStoklar(body);      break;
      case "efaturaOku":     result = manuelEfatura();      break;
      case "efatura2026":    result = manuelEfatura2026();    break;
      case "getDashboardData": result = getCombinedDashboardData(); break;
      default:               result = { error: "Bilinmeyen action: " + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    stokLogError(err);
    // ★ DÜZELTME: "erfr" yazım hatası vardı (tanımsız değişken), "err" olarak düzeltildi —
    // aksi halde bu catch bloğu kendisi de hata verip asıl hatayı gizliyordu.
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getAllData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const stokSheet = ss.getSheetByName(STOK_SHEETS.stoklar);
  let stoklar     = [];
  let stokTarih   = "";

  if (stokSheet) {
    const data = stokSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0].map(h => String(h).toUpperCase().trim());
      stoklar = data.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      }).filter(r => r.STOK_KODU);
      const tarihIdx = headers.indexOf("TARIH");
      if (tarihIdx >= 0) {
        for (let i = data.length - 1; i >= 1; i--) {
          const t  = data[i][tarihIdx];
          const ts = String(t);
          if (ts && !ts.includes("pasif")) {
            stokTarih = t instanceof Date
              ? Utilities.formatDate(t, Session.getScriptTimeZone(), "yyyy-MM-dd")
              : ts.split("T")[0];
            break;
          }
        }
      }
    }
  }

  let fiyatlar  = {};
  let fiyatGecmis = {};
  try {
    const sonuc = getFiyatlar();
    fiyatlar    = sonuc.fiyatMap;
    fiyatGecmis = sonuc.gecmisMap;
  } catch(e) {
    Logger.log("Fiyat okuma atlandı: " + e.message);
  }

  return {
    stoklar:     stoklar,
    stokTarih:   stokTarih,
    notlar:      sheetToObj(ss, STOK_SHEETS.notlar),
    etiketler:   sheetToArr(ss, STOK_SHEETS.etiketler),
    urunEt:      sheetToObj(ss, STOK_SHEETS.urunEt),
    markalar:    sheetToRows(ss, STOK_SHEETS.markalar),
    ayarlar:     sheetToObj(ss, STOK_SHEETS.ayarlar),
    fiyatlar:    fiyatlar,
    fiyatGecmis: fiyatGecmis,
  };
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const stokSheet = ss.getSheetByName(STOK_SHEETS.stoklar);
    const stokMap = {};
    let stokTarih = "";

    if (stokSheet && stokSheet.getLastRow() > 1) {
      const data = stokSheet.getDataRange().getValues();
      const h = data[0];
      const iKod = h.indexOf("STOK_KODU");
      const iAd = h.indexOf("STOK_ADI");
      const iDepo = h.indexOf("MERKEZ_DEPO");
      const iAyr = h.indexOf("AYRILMIS");
      const iSat = h.indexOf("SATILABILIR");
      const iTar = h.indexOf("TARIH");

      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        const kod = String(row[iKod] || "").trim();
        if (!kod) continue;
        stokMap[kod] = {
          stokAdi: iAd >= 0 ? String(row[iAd] || "") : "",
          merkezDepo: iDepo >= 0 ? Number(row[iDepo]) || 0 : 0,
          ayrilmis: iAyr >= 0 ? Number(row[iAyr]) || 0 : 0,
          satilabilir: iSat >= 0 ? Number(row[iSat]) || 0 : 0,
          stokTarih: iTar >= 0 ? String(row[iTar] || "") : ""
        };
        if (!stokTarih && iTar >= 0 && row[iTar] && !String(row[iTar]).includes("pasif")) {
          stokTarih = String(row[iTar]);
        }
      }
    }

    const fiyatSheet = ss.getSheetByName(SHEET_FIYAT);
    const fiyatMap = {};
    if (fiyatSheet && fiyatSheet.getLastRow() > 1) {
      const data = fiyatSheet.getDataRange().getValues();
      const h = data[0];
      const idx = {};
      h.forEach((val, i) => { idx[String(val)] = i; });
      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        const kod = String(row[idx['STOK_KODU']] || "").trim();
        if (!kod || fiyatMap[kod]) continue;
        fiyatMap[kod] = {
          birim: idx['BIRIM_FIYAT'] >= 0 ? row[idx['BIRIM_FIYAT']] || 0 : 0,
          iskonto: idx['ISKONTO'] >= 0 ? row[idx['ISKONTO']] || 0 : 0,
          net: idx['NET_FIYAT'] >= 0 ? row[idx['NET_FIYAT']] || 0 : 0,
          nakliye: idx['NAKLIYE_PAYI'] >= 0 ? row[idx['NAKLIYE_PAYI']] || 0 : 0,
          maliyet: idx['MALIYET_FIYAT'] >= 0 ? row[idx['MALIYET_FIYAT']] || 0 : 0,
          kdv: idx['KDV_ORANI'] >= 0 ? row[idx['KDV_ORANI']] || 0 : 0,
          faturaNo: idx['FATURA_NO'] >= 0 ? String(row[idx['FATURA_NO']] || "") : "",
          faturaTarih: idx['FATURA_TARIHI'] >= 0 ? String(row[idx['FATURA_TARIHI']] || "") : "",
          tedarikci: idx['TEDARIKCI'] >= 0 ? String(row[idx['TEDARIKCI']] || "") : "",
          faturaLink: idx['FATURA_LINK'] >= 0 ? String(row[idx['FATURA_LINK']] || "") : "",
          edmLink: idx['EDM_LINK'] >= 0 ? String(row[idx['EDM_LINK']] || "") : ""
        };
      }
    }

    const combined = [];
    Object.keys(stokMap).forEach(function(kod) {
      const s = stokMap[kod];
      const f = fiyatMap[kod] || {};
      combined.push({
        kod: kod, ad: s.stokAdi || f.stokAdi || "",
        merkezDepo: s.merkezDepo, ayrilmis: s.ayrilmis, satilabilir: s.satilabilir, stokTarih: s.stokTarih,
        birim: f.birim || 0, iskonto: f.iskonto || 0, net: f.net || 0,
        nakliye: f.nakliye || 0, maliyet: f.maliyet || 0, kdv: f.kdv || 0,
        faturaNo: f.faturaNo || "", faturaTarih: f.faturaTarih || "",
        tedarikci: f.tedarikci || "", faturaLink: f.faturaLink || "", edmLink: f.edmLink || ""
      });
    });
    Object.keys(fiyatMap).forEach(function(kod) {
      if (!stokMap[kod]) {
        const f = fiyatMap[kod];
        combined.push({
          kod: kod, ad: f.stokAdi || "", merkezDepo: 0, ayrilmis: 0, satilabilir: 0, stokTarih: "",
          birim: f.birim, iskonto: f.iskonto, net: f.net, nakliye: f.nakliye,
          maliyet: f.maliyet, kdv: f.kdv, faturaNo: f.faturaNo, faturaTarih: f.faturaTarih,
          tedarikci: f.tedarikci, faturaLink: f.faturaLink, edmLink: f.edmLink
        });
      }
    });

    let etiketler = [];
    let urunEt = {};
    let ayarlar = {};
    try {
      const etSheet = ss.getSheetByName(STOK_SHEETS.etiketler);
      if (etSheet) {
        etiketler = etSheet.getDataRange().getValues().slice(1)
          .filter(function(r){ return r[0]; })
          .map(function(r){ return { t: String(r[0]), c: String(r[1] || "#d97706") }; });
      }
      const ueSheet = ss.getSheetByName(STOK_SHEETS.urunEt);
      if (ueSheet) {
        ueSheet.getDataRange().getValues().slice(1).forEach(function(r){
          if (r[0]) urunEt[String(r[0])] = String(r[1] || "").split(",").filter(Boolean);
        });
      }
      const ayarSheet = ss.getSheetByName(STOK_SHEETS.ayarlar);
      if (ayarSheet) {
        ayarSheet.getDataRange().getValues().slice(1).forEach(function(r){
          if (r[0]) ayarlar[String(r[0])] = r[1];
        });
      }
    } catch(e) {}

    combined.forEach(function(item){
      item.etiketler = urunEt[item.kod] || [];
    });

    return { data: combined, stokTarih: stokTarih, etiketler: etiketler, urunEt: urunEt, ayarlar: ayarlar };
  } catch (e) {
    Logger.log("getDashboardData hata: " + e.message);
    return { error: e.message, data: [] };
  }
}
function saveNot(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, STOK_SHEETS.notlar, ["STOK_KODU","NOT","TARIH"]);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(body.kod)) {
      if (body.not) sheet.getRange(i+1,2,1,2).setValues([[body.not, new Date()]]);
      else sheet.deleteRow(i+1);
      return { ok: true };
    }
  }
  if (body.not) sheet.appendRow([body.kod, body.not, new Date()]);
  return { ok: true };
}

function saveEtiket(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, STOK_SHEETS.etiketler, ["ETIKET","RENK"]);
  sheet.clearContents();
  sheet.appendRow(["ETIKET","RENK"]);
  (body.etiketler || []).forEach(e => sheet.appendRow([e.t||e, e.c||"#d97706"]));
  return { ok: true };
}

function saveUrunEt(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, STOK_SHEETS.urunEt, ["STOK_KODU","ETIKETLER"]);
  sheet.clearContents();
  sheet.appendRow(["STOK_KODU","ETIKETLER"]);
  Object.entries(body.urunEt || {}).forEach(([kod, tags]) => {
    if (tags && tags.length > 0) sheet.appendRow([kod, tags.join(",")]);
  });
  return { ok: true };
}

function saveMarka(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, STOK_SHEETS.markalar, ["KOD","AD","RENK"]);
  sheet.clearContents();
  sheet.appendRow(["KOD","AD","RENK"]);
  (body.markalar || []).forEach(m => sheet.appendRow([m.k, m.a, m.c||"#2563eb"]));
  return { ok: true };
}

function saveAyar(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, STOK_SHEETS.ayarlar, ["ANAHTAR","DEGER"]);
  sheet.clearContents();
  sheet.appendRow(["ANAHTAR","DEGER"]);
  Object.entries(body.ayarlar || {}).forEach(([k,v]) => sheet.appendRow([k, v]));
  return { ok: true };
}

function sheetToObj(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  data.slice(1).forEach(row => { if (row[0]) result[String(row[0])] = row[1]; });
  return result;
}

function sheetToArr(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .map(row => ({ t: row[0], c: row[1] || "#d97706" }))
    .filter(e => e.t);
}

function sheetToRows(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .map(row => ({ k: String(row[0]), a: row[1], c: row[2] || "#2563eb" }))
    .filter(m => m.k);
}

// ═══════════════════════════════════════════════════════════════════
//  TETİKLEYİCİLER
// ═══════════════════════════════════════════════════════════════════

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("efaturaOku").timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger("checkStokMail").timeBased().everyMinutes(15).create();
  Logger.log("✅ Tetikleyiciler kuruldu: e-fatura + stok (fincanlaryapi)");
}

// ═══════════════════════════════════════════════════════════════════
//  TEST
// ═══════════════════════════════════════════════════════════════════

function manuelTest() {
  Logger.log("=== E-FATURA MANUEL TEST ===");
  efaturaOku();
  Logger.log("=== BİTTİ ===");
}

function stokTest() {
  Logger.log("=== STOK MAİL MANUEL TEST ===");
  checkStokMail();
  Logger.log("=== BİTTİ ===");
}

function fiyatTest() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_FIYAT);
  if (!sh) { Logger.log("FATURAFIYAT bulunamadı"); return; }
  var data = sh.getDataRange().getValues();
  Logger.log("Toplam kayıt: " + (data.length - 1));
  Logger.log("Header: " + data[0].join(" | "));
  if (data.length > 1) Logger.log("İlk kayıt: " + data[1].join(" | "));
}

function logKontrol() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) { Logger.log("LOG SHEET YOK!"); return; }
  var data = sh.getDataRange().getValues();
  Logger.log("LOG satır sayısı: " + data.length);
  if (data.length > 1) Logger.log("Son kayıt: " + JSON.stringify(data[data.length-1]));
}

function pdfLinkTest() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_FIYAT);
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const iLink = h.indexOf("FATURA_LINK");
  const iEdm = h.indexOf("EDM_LINK");
  const iFNo = h.indexOf("FATURA_NO");

  Logger.log("Son 5 fatura link kontrolü:");
  for (let r = Math.max(1, data.length - 5); r < data.length; r++) {
    const link = data[r][iLink] || "";
    const edm = data[r][iEdm] || "";
    const fno = data[r][iFNo] || "";
    Logger.log(fno + " | PDF: " + (link ? "VAR" : "YOK") + " | EDM: " + (edm ? "VAR" : "YOK"));
  }
}

function eskiDriveLinkDuzelt() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_FIYAT);
  var data = sh.getDataRange().getValues();
  var h = data[0];
  var iLink = h.indexOf("FATURA_LINK");

  if (iLink === -1) { Logger.log("FATURA_LINK bulunamadı"); return; }

  var guncellenen = 0;
  for (var r = 1; r < data.length; r++) {
    var url = String(data[r][iLink] || "");
    if (url.indexOf("drive.google.com/file/d/") > -1 && url.indexOf("/preview") === -1) {
      var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (m) {
        data[r][iLink] = "https://drive.google.com/file/d/" + m[1] + "/preview";
        guncellenen++;
      }
    }
  }

  if (guncellenen > 0) {
    sh.getDataRange().setValues(data);
    Logger.log("✅ " + guncellenen + " eski Drive linki /preview formatına çevrildi.");
  } else {
    Logger.log("Düzeltilecek eski link bulunamadı.");
  }
}

function faturaTarihiDuzelt(faturaNo, yeniTarih) {
  // yeniTarih format: "dd/MM/yyyy" (örn: "18/05/2026")
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_FIYAT);
  var data = sh.getDataRange().getValues();
  var h = data[0];
  var iFNo = h.indexOf("FATURA_NO");
  var iFTar = h.indexOf("FATURA_TARIHI");

  sh.getRange(2, iFTar + 1, sh.getLastRow() - 1, 1).setNumberFormat("@");

  var guncellenen = 0;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iFNo] || "") === faturaNo) {
      data[r][iFTar] = yeniTarih;
      guncellenen++;
    }
  }

  if (guncellenen > 0) {
    sh.getDataRange().setValues(data);
    Logger.log("✅ " + faturaNo + " → " + yeniTarih + " (" + guncellenen + " satır)");
  } else {
    Logger.log("Fatura bulunamadı: " + faturaNo);
  }
}

function faturaBulTest() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_FIYAT);
  var data = sh.getDataRange().getValues();
  var h = data[0];
  var iFNo = h.indexOf("FATURA_NO");
  var iFTar = h.indexOf("FATURA_TARIHI");

  var aranan = "CNY2026000001303";
  Logger.log("Aranan: [" + aranan + "]");
  Logger.log("Toplam satır: " + data.length);
  Logger.log("FATURA_NO kolon index: " + iFNo);

  var bulunan = 0;
  for (var r = 1; r < Math.min(data.length, 20); r++) {
    var val = String(data[r][iFNo] || "");
    if (val.indexOf("1303") > -1) {
      Logger.log("Satır " + (r+1) + ": [" + val + "] | Tarih: [" + data[r][iFTar] + "]");
      bulunan++;
    }
  }
  Logger.log("Bulunan: " + bulunan);
}

function hataliTarihBul() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_FIYAT);
  var data = sh.getDataRange().getValues();
  var h = data[0];
  var iFNo = h.indexOf("FATURA_NO");
  var iFTar = h.indexOf("FATURA_TARIHI");
  var iTed = h.indexOf("TEDARIKCI");

  Logger.log("=== ŞÜPHELİ TARİHLER ===");
  var supheli = 0;

  for (var r = 1; r < data.length; r++) {
    var fno = String(data[r][iFNo] || "");
    var ftarRaw = data[r][iFTar];
    var ftar = (ftarRaw instanceof Date)
      ? Utilities.formatDate(ftarRaw, "Europe/Istanbul", "dd/MM/yyyy")
      : String(ftarRaw || "");
    var ted = String(data[r][iTed] || "");

    if (ted !== "CANYAP") continue;

    var m = ftar.match(/(\d{2})\/(\d{2})\/(\d{4})/);
    if (!m) continue;

    var faturaYil = parseInt(m[3]);
    var faturaAy = parseInt(m[2]);

    var noAy = fno.match(/CNY(\d{4})(\d{2})/);
    if (noAy) {
      var noYil = parseInt(noAy[1]);
      var noAyNum = parseInt(noAy[2]);

      if (noYil !== faturaYil || Math.abs(noAyNum - faturaAy) > 2) {
        Logger.log("Şüpheli: " + fno + " | Fatura Tarihi: " + ftar + " | No'daki ay/yıl: " + noYil + "-" + noAyNum);
        supheli++;
      }
    }
  }

  Logger.log("Toplam şüpheli: " + supheli);
}

// ═══════════════════════════════════════════════════════════════════
//  ★ TARİH ONARIMI #2 — "Vade Tarihi" ile karışmış olan FATURA_TARIHI
//  değerlerini, faturanın kendi sayfasına tekrar gidip DOĞRU tarihi
//  (tarihCikarHTMLden'in düzeltilmiş mantığıyla) yeniden çekerek düzeltir.
//  Menüden (Çalıştır ▶) tarihleriYenidenCek fonksiyonunu bir kez çalıştırın.
//  Not: EDM_LINK sütununun kalıcı (ViewInvoice) link içermesi gerekir —
//  önce linkVeTarihOnar() (Onarım #1, aşağıda) çalıştırılmış olmalı.
// ═══════════════════════════════════════════════════════════════════
function tarihleriYenidenCek() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_FIYAT);
  if (!sh) { Logger.log("FATURAFIYAT bulunamadı"); return; }

  var data = sh.getDataRange().getValues();
  if (data.length < 2) { Logger.log("Veri yok"); return; }

  var h = data[0];
  var iFNo  = h.indexOf("FATURA_NO");
  var iFTar = h.indexOf("FATURA_TARIHI");
  var iEdm  = h.indexOf("EDM_LINK");

  if (iFNo === -1 || iFTar === -1 || iEdm === -1) {
    Logger.log("Gerekli kolonlar bulunamadı");
    return;
  }

  // Her fatura numarası için EDM linkini bir kez topla (satırlar tekrarlı olabilir)
  var faturaToEdm = {};
  for (var r = 1; r < data.length; r++) {
    var fno = String(data[r][iFNo] || "").trim();
    var edm = String(data[r][iEdm] || "").trim();
    if (fno && edm && !faturaToEdm[fno]) faturaToEdm[fno] = edm;
  }

  var faturaToYeniTarih = {};
  var basarili = 0, hatali = 0;
  for (var fno2 in faturaToEdm) {
    var link = faturaToEdm[fno2];
    if (link.indexOf("tmp-ef") > -1) {
      Logger.log("Atlandı (link hâlâ geçici, önce linkVeTarihOnar çalıştırın): " + fno2);
      hatali++;
      continue;
    }
    try {
      var anaHTML = sayfaCek(link);
      if (!anaHTML) throw new Error("Ana sayfa açılamadı");
      var htmlLink = htmlLinkBul(anaHTML);
      if (!htmlLink) throw new Error("Fatura HTML linki bulunamadı");
      var faturaHTML = sayfaCek(htmlLink);
      if (!faturaHTML) throw new Error("Fatura HTML açılamadı");

      var yeniTarih = tarihCikarHTMLden(faturaHTML, new Date());
      faturaToYeniTarih[fno2] = yeniTarih;
      Logger.log(fno2 + " → " + yeniTarih);
      basarili++;
    } catch(e) {
      Logger.log("HATA (" + fno2 + "): " + e.message);
      hatali++;
    }
    Utilities.sleep(400); // rate limit koruması
  }

  // Tüm satırları güncelle
  var guncellenenSatir = 0;
  for (var r2 = 1; r2 < data.length; r2++) {
    var fno3 = String(data[r2][iFNo] || "").trim();
    if (faturaToYeniTarih[fno3]) {
      data[r2][iFTar] = faturaToYeniTarih[fno3];
      guncellenenSatir++;
    }
  }

  sh.getRange(2, iFTar + 1, sh.getLastRow() - 1, 1).setNumberFormat("@");
  sh.getDataRange().setValues(data);

  Logger.log("✅ TARİH ONARIMI TAMAMLANDI");
  Logger.log("Başarıyla yeniden çekilen fatura: " + basarili);
  Logger.log("Atlanan/hatalı fatura: " + hatali);
  Logger.log("Güncellenen satır sayısı: " + guncellenenSatir);
}

// ═══════════════════════════════════════════════════════════════════
//  ★ TEK SEFERLİK ONARIM #1 — mevcut bozuk EDM linklerini (geçici→kalıcı)
//  ve Date-nesnesi olarak bozulmuş tarih hücrelerini (görüntü sorunu) düzeltir.
//  Menüden (Çalıştır ▶) linkVeTarihOnar fonksiyonunu ÖNCE bir kez çalıştırın,
//  ardından tarihleriYenidenCek() (Onarım #2, yukarıda) çalıştırın.
//  Kalıcı hasar yok: sadece FATURA_TARIHI ve EDM_LINK sütunlarına dokunur.
// ═══════════════════════════════════════════════════════════════════
function linkVeTarihOnar() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_FIYAT);
  if (!sh) { Logger.log("FATURAFIYAT bulunamadı"); return; }

  var data = sh.getDataRange().getValues();
  if (data.length < 2) { Logger.log("Veri yok"); return; }

  var h = data[0];
  var iFTar = h.indexOf("FATURA_TARIHI");
  var iEdm  = h.indexOf("EDM_LINK");

  if (iFTar === -1 || iEdm === -1) {
    Logger.log("FATURA_TARIHI veya EDM_LINK kolonu bulunamadı");
    return;
  }

  // ── 1) TARİHLERİ ONAR: gerçek Date nesnelerini "dd/MM/yyyy" metnine çevir ──
  var tarihDuzeltildi = 0;
  for (var r = 1; r < data.length; r++) {
    var v = data[r][iFTar];
    if (v instanceof Date) {
      data[r][iFTar] = Utilities.formatDate(v, "Europe/Istanbul", "dd/MM/yyyy");
      tarihDuzeltildi++;
    }
  }
  // Sütunu kalıcı olarak düz metne zorla — bir daha otomatik tarihe dönüşmesin
  sh.getRange(2, iFTar + 1, sh.getLastRow() - 1, 1).setNumberFormat("@");

  // ── 2) LİNKLERİ ONAR: "tmp-ef" (geçici) linkleri kalıcı ViewInvoice linkine çevir ──
  // NOT: Buradaki "2030011228" ön eki, örnek faturalarda sabit görünen bir
  // EDM Bilişim portal/şirket kimliğidir. Onarımdan sonra 2-3 faturayı elle
  // test edin; eğer açılmıyorsa bu ön ek yanlış demektir, haber verin.
  var EDM_PORTAL_ID = "2030011228";
  var linkDuzeltildi = 0, linkAtlandi = 0;
  for (var r2 = 1; r2 < data.length; r2++) {
    var edm = String(data[r2][iEdm] || "");
    if (edm.indexOf("tmp-ef") === -1) continue; // zaten kalıcı link, dokunma
    var m = edm.match(/tmp-ef\/([a-f0-9-]{36})/i);
    if (!m) { linkAtlandi++; continue; }
    var uuid = m[1];
    data[r2][iEdm] = EDM_BASE + "/fatura/ViewInvoice/" + EDM_PORTAL_ID + "/" + uuid + "/efatura";
    linkDuzeltildi++;
  }

  sh.getDataRange().setValues(data);

  Logger.log("✅ ONARIM TAMAMLANDI");
  Logger.log("Tarih düzeltilen satır: " + tarihDuzeltildi);
  Logger.log("Link düzeltilen satır: " + linkDuzeltildi);
  Logger.log("Link düzeltilemeyen (UUID bulunamadı) satır: " + linkAtlandi);
}

// Fatura HTML'inden sadece STOK_KODU → MIKTAR eşleşmesini çıkarır
// (faturaParseEt'in tablo ayrıştırma mantığının küçültülmüş hali).
function miktarlariCikarHTMLden(html) {
  var sonuc = {};
  var temizHTML = html;
  while (temizHTML.indexOf("<script") > -1) {
    var s = temizHTML.toLowerCase().indexOf("<script");
    var e = temizHTML.toLowerCase().indexOf("</script>", s);
    if (e === -1) break;
    temizHTML = temizHTML.substring(0, s) + temizHTML.substring(e + 9);
  }
  while (temizHTML.indexOf("<style") > -1) {
    var s2 = temizHTML.toLowerCase().indexOf("<style");
    var e2 = temizHTML.toLowerCase().indexOf("</style>", s2);
    if (e2 === -1) break;
    temizHTML = temizHTML.substring(0, s2) + temizHTML.substring(e2 + 8);
  }

  function sayiyaCevir(str) {
    if (!str) return 0;
    // ★ DÜZELTME (bkz. faturaParseEt içindeki aynı isimli fonksiyon): sadece string
    // içindeki İLK bitişik sayısal belirteç alınıyor — "38,88 (2)" gibi bir hücrede
    // parantez içindeki koli sayısı artık asıl miktara karışıp "38,882" üretmiyor.
    var m = String(str).match(/\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+(?:,\d+)?/);
    if (!m) return 0;
    var t = m[0];
    t = t.replace(/\.(?=\d{3}(?:[,]|$))/g, "");
    t = t.replace(",", ".");
    return parseFloat(t) || 0;
  }

  var trPat = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  var trM;
  while ((trM = trPat.exec(temizHTML)) !== null) {
    var satir = trM[1];
    var tdler = [];
    var tdPat = /<td[^>]*>([\s\S]*?)<\/td>/gi;
    var tdM;
    while ((tdM = tdPat.exec(satir)) !== null) {
      var ic = tdM[1].replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/\s+/g, " ").trim();
      if (ic) tdler.push(ic);
    }
    if (tdler.length < 3) continue;

    var stokKoduIdx = -1, stokKodu = "";
    for (var i = 0; i < tdler.length; i++) {
      if (/^\d{8,15}$/.test(tdler[i].trim())) { stokKodu = tdler[i].trim(); stokKoduIdx = i; break; }
    }
    if (!stokKodu) continue;

    var miktar = sayiyaCevir(tdler[stokKoduIdx + 2] || "");
    if (miktar > 0) sonuc[stokKodu] = miktar;
  }
  return sonuc;
}

// ★ TEK SEFERLİK ONARIM #3 — MIKTAR kolonu eklenmeden önce yazılmış eski
// kayıtları, EDM_LINK üzerinden faturayı yeniden çekip miktarla geriye dönük
// doldurur. Sadece MIKTAR'ı 0/boş olan satırları işler, diğer hiçbir alana
// (fiyat, iskonto, tarih vb.) dokunmaz. Her fatura işlendikçe İLGİLİ SATIRLAR
// ANINDA sayfaya yazılır (tüm işi hafızada tutup en sonda tek seferde yazmaz) —
// bu yüzden "Exceeded maximum execution time" hatası alsanız bile o ana kadar
// işlenen faturalar KAYBOLMAZ. 5 dakika dolunca kendini güvenli şekilde durdurur;
// fonksiyonu tekrar çalıştırmanız kaldığı yerden devam etmesi için yeterli.
function miktarlariGeriDoldur() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_FIYAT);
  if (!sh) { Logger.log("FATURAFIYAT bulunamadı"); return; }

  var data = sh.getDataRange().getValues();
  if (data.length < 2) { Logger.log("Veri yok"); return; }

  var h = data[0];
  var iFNo = h.indexOf("FATURA_NO");
  var iEdm = h.indexOf("EDM_LINK");
  var iMik = h.indexOf("MIKTAR");
  var iKod = h.indexOf("STOK_KODU");

  if (iFNo === -1 || iEdm === -1 || iMik === -1 || iKod === -1) {
    Logger.log("Gerekli kolonlar bulunamadı (FATURA_NO/EDM_LINK/MIKTAR/STOK_KODU)");
    return;
  }

  // Fatura no → { edmLink, rows: [satır indeksleri] } — sadece MIKTAR'ı boş satırlar.
  var faturaMap = {};
  for (var r = 1; r < data.length; r++) {
    var mevcutMiktar = parseFloat(data[r][iMik]) || 0;
    if (mevcutMiktar > 0) continue;
    var fno = String(data[r][iFNo] || "").trim();
    if (!fno) continue;
    if (!faturaMap[fno]) faturaMap[fno] = { edmLink: String(data[r][iEdm] || "").trim(), rows: [] };
    faturaMap[fno].rows.push(r);
  }

  var faturaNolar = Object.keys(faturaMap);
  if (faturaNolar.length === 0) { Logger.log("Miktarı boş satır yok, yapılacak iş yok."); return; }
  Logger.log("İşlenecek fatura sayısı: " + faturaNolar.length);

  var baslangic = new Date().getTime();
  var SURE_LIMITI_MS = 5 * 60 * 1000; // 6 dk sert limitin altında güvenli pay
  var basarili = 0, hatali = 0, atlanan = 0, guncellenenSatir = 0, durduruldu = false;

  for (var i = 0; i < faturaNolar.length; i++) {
    if (new Date().getTime() - baslangic > SURE_LIMITI_MS) {
      Logger.log("⏱ Süre limiti yaklaştı, güvenli şekilde durduruldu (" + i + "/" + faturaNolar.length + " işlendi). Fonksiyonu tekrar çalıştırın, kaldığı yerden devam eder.");
      durduruldu = true;
      break;
    }
    var fno = faturaNolar[i];
    var bilgi = faturaMap[fno];
    var link = bilgi.edmLink;

    if (!link || link.indexOf("tmp-ef") > -1) {
      Logger.log("Atlandı (link yok/geçici, önce linkVeTarihOnar çalıştırın): " + fno);
      atlanan++;
      continue;
    }

    try {
      var anaHTML = sayfaCek(link);
      if (!anaHTML) throw new Error("Ana sayfa açılamadı");
      var htmlLink = htmlLinkBul(anaHTML);
      if (!htmlLink) throw new Error("Fatura HTML linki bulunamadı");
      var faturaHTML = sayfaCek(htmlLink);
      if (!faturaHTML) throw new Error("Fatura HTML açılamadı");

      var miktarMap = miktarlariCikarHTMLden(faturaHTML);

      // Bu faturanın satırlarını HEMEN sayfaya yaz — kesinti olsa da bu kayıp gitmez.
      var buFaturaGuncellendi = 0;
      for (var j = 0; j < bilgi.rows.length; j++) {
        var rIdx = bilgi.rows[j];
        var kod = String(data[rIdx][iKod] || "").trim();
        if (miktarMap[kod] > 0) {
          sh.getRange(rIdx + 1, iMik + 1).setValue(miktarMap[kod]);
          buFaturaGuncellendi++;
        }
      }
      guncellenenSatir += buFaturaGuncellendi;
      Logger.log(fno + " → " + buFaturaGuncellendi + "/" + bilgi.rows.length + " satır güncellendi");
      basarili++;
    } catch(e) {
      Logger.log("HATA (" + fno + "): " + e.message);
      hatali++;
    }
    Utilities.sleep(300);
  }

  Logger.log(durduruldu ? "⏱ BU ÇALIŞTIRMA SÜRE LİMİTİNDEN DURDU (devam etmek için tekrar çalıştırın)" : "✅ TÜM FATURALAR TAMAMLANDI");
  Logger.log("Bu çalıştırmada: " + basarili + " başarılı, " + hatali + " hatalı, " + atlanan + " atlandı");
  Logger.log("Güncellenen satır sayısı: " + guncellenenSatir);
}
