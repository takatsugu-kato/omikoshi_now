// 管理者パスワードを保存するスクリプトプロパティのキー
var ADMIN_PW_KEY = 'ADMIN_PASSWORD';
// ログインセッションの有効時間（秒）＝12時間
var SESSION_TTL_SEC = 12 * 60 * 60;
// Google Maps API キーを保存するスクリプトプロパティのキー
var GMAP_KEY = 'GMAP_API_KEY';

function doGet(e) {
  // パラメータ mode を取得（小文字に統一）
  var mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode.toLowerCase() : "";
  var token = (e && e.parameter && e.parameter.token) ? e.parameter.token : "";
  
  // ?mode=admin のときだけ Index（役員用操作画面）を表示（要ログイン）
  if (mode === 'admin') {
    // セッショントークンを検証
    if (token && CacheService.getScriptCache().get('admin_session_' + token)) {
      return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('神輿なう - 管理')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    // 未ログインならログイン画面を表示
    return HtmlService.createHtmlOutputFromFile('Login')
      .setTitle('神輿なう - ログイン')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // それ以外（パラメータなし含む）はすべて Viewer（閲覧画面）を表示
  return HtmlService.createHtmlOutputFromFile('Viewer')
    .setTitle('神輿なう')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ログイン処理：パスワードを検証してセッショントークンを発行する
function login(password) {
  var stored = getAdminPassword();
  if (!stored) {
    return { success: false, message: "管理者パスワードが未設定です（スクリプトプロパティ ADMIN_PASSWORD を設定してください）" };
  }
  if (password === stored) {
    var token = Utilities.getUuid();
    CacheService.getScriptCache().put('admin_session_' + token, '1', SESSION_TTL_SEC);
    // 正しい Web アプリ URL（exec）を返す。iframe 内の location は userCodeAppPanel なので使えない
    return { success: true, token: token, baseUrl: ScriptApp.getService().getUrl() };
  }
  return { success: false, message: "パスワードが違います" };
}

// 管理者パスワードをスクリプトプロパティから取得する
function getAdminPassword() {
  var pw = PropertiesService.getScriptProperties().getProperty(ADMIN_PW_KEY);
  return pw ? pw : "";
}

// Google Maps API キーをスクリプトプロパティから取得する（コードにベタ書きしない）
function getGmapApiKey() {
  var key = PropertiesService.getScriptProperties().getProperty(GMAP_KEY);
  return key ? key : "";
}

// ショートカットアプリ等からのPOSTリクエストを受信
function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
    var lat = Number(contents.lat);
    var lng = Number(contents.lng);
    
    if (lat && lng) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var gpsSheet = ss.getSheetByName('GPSログ');
      
      // GEO情報の作成（Google Mapsリンク文字列例: "35.38858,139.403229" または リンクURL）
      var geo = lat + ',' + lng; 
      // もし Google Maps の URL 形式にしたい場合は以下を有効化してください
      // var geo = "https://www.google.com/maps?q=" + lat + "," + lng;

      // GPSログシートに [タイムスタンプ, 緯度, 経度, GEO] を追記
      gpsSheet.appendRow([new Date(), lat, lng, geo]);
      
      return ContentService.createTextOutput(JSON.stringify({ status: "success", lat: lat, lng: lng, geo: geo }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "lat or lng is missing" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ★追加：休憩所マスターのB列からリストを取得し、記録済み状態（到着/出発）も付与する関数
function getStations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('休憩所マスター');
  var stations = [];
  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      // A列=ID, B列=名前
      var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      values.forEach(function(row) {
        if (row[1]) stations.push({ name: row[1], id: row[0] });
      });
    }
  }

  // 休憩所記録から到着・出発の記録済み状態を取得して付与
  var historySheet = ss.getSheetByName('休憩所記録');
  var doneMap = {};
  if (historySheet && historySheet.getLastRow() >= 2) {
    var hValues = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 3).getValues();
    hValues.forEach(function(row) {
      var name = row[1];
      var type = row[2];
      if (!doneMap[name]) doneMap[name] = {};
      if (type === "到着" || type === "出発") doneMap[name][type] = true;
    });
  }
  stations.forEach(function(s) {
    s.arrived = !!(doneMap[s.name] && doneMap[s.name]["到着"]);
    s.departed = !!(doneMap[s.name] && doneMap[s.name]["出発"]);
  });

  return stations;
}

// 手動ボタン（到着・出発）の記録用関数
// 記録構成: A列=記録日時(タイムスタンプ), B列=休憩所名, C列=到着/出発, D列=コメント, E列=未使用
function recordStation(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('休憩所記録');
  var now = new Date();
  sheet.appendRow([
    now,              // A列: 記録日時（タイムスタンプ）
    data.location,    // B列: 休憩所名
    data.status,      // C列: 到着 or 出発
    data.comment || "", // D列: コメント
    ""                // E列: 未使用
  ]);
  return "【" + data.location + "】" + data.status + " を記録しました";
}

// 自動GPS送信の記録用関数
function recordGPS(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('GPSログ');
  
  // 緯度と経度をカンマで結合（例: "35.3910945717178,139.408267572967"）
  var geoLocation = data.latitude + "," + data.longitude;
  
  sheet.appendRow([
    new Date(),
    data.latitude,
    data.longitude,
    geoLocation // D列に結合した文字列を書き込み
  ]);
  return "GPS更新完了";
}


function getViewerData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 最新GPS
  var gpsSheet = ss.getSheetByName('GPSログ');
  var latestGps = null;
  if (gpsSheet && gpsSheet.getLastRow() >= 2) {
    var lastRow = gpsSheet.getLastRow();
    var gpsValues = gpsSheet.getRange(lastRow, 1, 1, 3).getValues()[0];
    latestGps = {
      timestamp: Utilities.formatDate(new Date(gpsValues[0]), "JST", "HH:mm"),
      lat: Number(gpsValues[1]),
      lng: Number(gpsValues[2])
    };
  }
  
  // 2. 休憩所マスターの全リスト取得
  var masterSheet = ss.getSheetByName('休憩所マスター');
  var masterList = [];
  if (masterSheet && masterSheet.getLastRow() >= 2) {
    // 列構成: A=id, B=name, C=target_arrival, D=target_departure, E=latitude, F=longitude
    var values = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 6).getValues();
    values.forEach(function(row) {
      var rawLat = row[4]; // E列: latitude
      var rawLng = row[5]; // F列: longitude
      var lat = (rawLat === "" || rawLat === null || rawLat === undefined || isNaN(Number(rawLat))) ? null : Number(rawLat);
      var lng = (rawLng === "" || rawLng === null || rawLng === undefined || isNaN(Number(rawLng))) ? null : Number(rawLng);
      masterList.push({
        id: row[0],
        name: row[1],
        target_arrival: row[2] ? formatTimeStr(row[2]) : "",
        target_departure: row[3] ? formatTimeStr(row[3]) : "",
        lat: lat,
        lng: lng
      });
    });
  }
  
  // 2-2. トイレマスターの全リスト取得（A=id, B=name, C=latitude, D=longitude）
  var toiletSheet = ss.getSheetByName('トイレマスター');
  var toilets = [];
  if (toiletSheet && toiletSheet.getLastRow() >= 2) {
    var tValues = toiletSheet.getRange(2, 1, toiletSheet.getLastRow() - 1, 4).getValues();
    tValues.forEach(function(row) {
      var rawLat = row[2]; // C列: latitude
      var rawLng = row[3]; // D列: longitude
      var lat = (rawLat === "" || rawLat === null || rawLat === undefined || isNaN(Number(rawLat))) ? null : Number(rawLat);
      var lng = (rawLng === "" || rawLng === null || rawLng === undefined || isNaN(Number(rawLng))) ? null : Number(rawLng);
      if (row[1]) {
        toilets.push({
          id: row[0],
          name: row[1],
          lat: lat,
          lng: lng
        });
      }
    });
  }
  
  // 3. 休憩所記録の取得（全記録）
  var historySheet = ss.getSheetByName('休憩所記録');
  var historyMap = {}; // name -> { arrival: Date, departure: Date }
  var lastRecord = null;
  
  if (historySheet && historySheet.getLastRow() >= 2) {
    // A列: 記録日時(タイムスタンプ), B列: 休憩所名, C列: 到着/出発, D列: コメント, E列: (未使用)
    var hValues = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 5).getValues();
    hValues.forEach(function(row) {
      var time = new Date(row[0]); // タイムスタンプはA列の値を使用
      var name = row[1];
      var type = row[2]; // "到着" or "出発"
      var comment = row[3] ? String(row[3]) : "";
      
      if (!historyMap[name]) historyMap[name] = {};
      historyMap[name][type] = time;
      historyMap[name][type + "コメント"] = comment;
      
      lastRecord = { name: name, type: type, time: time };
    });
  }
  
  // 4. ステータスとタイムラインの組み立て
  var timeline = [];
  var currentSpot = null;
  var nextSpot = null;
  var isStag = false; // 滞在中か移動中か
  var currentIdx = -1;
  
  // 現在地の判定
  if (lastRecord) {
    for (var i = 0; i < masterList.length; i++) {
      if (masterList[i].name === lastRecord.name) {
        currentIdx = i;
        if (lastRecord.type === "到着") {
          isStag = true;
          currentSpot = masterList[i];
          nextSpot = masterList[i + 1] || null;
        } else {
          // 出発済みの場合は「次の休憩所」へ移動中
          isStag = false;
          currentSpot = masterList[i]; // 直近出た場所
          nextSpot = masterList[i + 1] || null;
        }
        break;
      }
    }
  } else if (masterList.length > 0) {
    nextSpot = masterList[0]; // まだ始まっていない場合
  }
  
  // 遅延時間の計算（直近の記録時刻と予定時刻の差で固定）
  // ※ 現在時刻ではなく「記録を打った時刻」を基準にするため、
  //    休憩中・移動中でも刻々と増え続けることはない
  var delayMinutes = 0;
  
  if (lastRecord) {
    // 直近の記録に対応する休憩所の予定時刻を探す
    var recSpot = null;
    for (var i = 0; i < masterList.length; i++) {
      if (masterList[i].name === lastRecord.name) {
        recSpot = masterList[i];
        break;
      }
    }
    if (recSpot) {
      var p;
      var target = new Date(lastRecord.time);
      if (lastRecord.type === "到着" && recSpot.target_arrival) {
        // 到着の遅延 = 実際到着時刻 − 予定到着時刻
        p = recSpot.target_arrival.split(":");
        target.setHours(parseInt(p[0], 10), parseInt(p[1], 10), 0, 0);
        delayMinutes = Math.round((lastRecord.time - target) / (1000 * 60));
      } else if (lastRecord.type === "出発" && recSpot.target_departure) {
        // 出発の遅延 = 実際出発時刻 − 予定出発時刻
        p = recSpot.target_departure.split(":");
        target.setHours(parseInt(p[0], 10), parseInt(p[1], 10), 0, 0);
        delayMinutes = Math.round((lastRecord.time - target) / (1000 * 60));
      }
    }
  }

  // タイムラインのバッジ判定
  masterList.forEach(function(spot, index) {
    var statusText = "予定";
    var badgeClass = "badge-future";
    
    if (historyMap[spot.name]) {
      if (historyMap[spot.name]["到着"] && historyMap[spot.name]["出発"]) {
        statusText = "通過済";
        badgeClass = "badge-passed";
      } else if (historyMap[spot.name]["到着"]) {
        statusText = "滞在中";
        badgeClass = "badge-stay";
      }
    } else if (nextSpot && spot.name === nextSpot.name && !isStag) {
      statusText = "次の休憩所";
      badgeClass = "badge-next";
    }
    
    var arr  = historyMap[spot.name] ? historyMap[spot.name]["到着"] : null;
    var dep  = historyMap[spot.name] ? historyMap[spot.name]["出発"] : null;
    var cArr = historyMap[spot.name] ? historyMap[spot.name]["到着コメント"] : "";
    var cDep = historyMap[spot.name] ? historyMap[spot.name]["出発コメント"] : "";

    timeline.push({
      id: spot.id,
      name: spot.name,
      arrival: spot.target_arrival,
      departure: spot.target_departure,
      actual_arrival: arr ? Utilities.formatDate(arr, "JST", "HH:mm") : "",
      actual_departure: dep ? Utilities.formatDate(dep, "JST", "HH:mm") : "",
      arrival_comment: cArr,
      departure_comment: cDep,
      statusText: statusText,
      badgeClass: badgeClass,
      lat: spot.lat,
      lng: spot.lng
    });
  });

  // 現在の休憩所（滞在中）の到着予定・実際到着時刻を算出
  var currentArrival = currentSpot ? currentSpot.target_arrival : "";
  var currentDeparture = currentSpot ? currentSpot.target_departure : "";
  var actualArrival = "";
  var actualDeparture = "";
  if (currentSpot && historyMap[currentSpot.name]) {
    if (historyMap[currentSpot.name]["到着"]) {
      actualArrival = Utilities.formatDate(historyMap[currentSpot.name]["到着"], "JST", "HH:mm");
    }
    if (historyMap[currentSpot.name]["出発"]) {
      actualDeparture = Utilities.formatDate(historyMap[currentSpot.name]["出発"], "JST", "HH:mm");
    }
  }

  return {
    gps: latestGps,
    statusCard: {
      isStag: isStag,
      currentName: currentSpot ? currentSpot.name : "",
      nextName: nextSpot ? nextSpot.name : "",
      nextTime: nextSpot ? nextSpot.target_arrival : "",
      currentArrival: currentArrival,
      currentDeparture: currentDeparture,
      actualArrival: actualArrival,
      actualDeparture: actualDeparture,
      currentIndex: currentIdx + 1,
      delayMinutes: delayMinutes
    },
    timeline: timeline,
    toilets: toilets
  };
}

function formatTimeStr(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "JST", "HH:mm");
  }
  return String(val);
}
