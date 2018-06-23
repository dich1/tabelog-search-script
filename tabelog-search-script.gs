function onOpen() {
  // メニューバーにカスタムメニューを追加
  //var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //var entries = [
  //  {name : "情報を取得"  , functionName : "myFunction"},
  //];
  //spreadsheet.addMenu("追加メニュー", entries);
  deleteTrigger();
  deleteProperty("add_sheet");
  deleteProperty("cur_page_url");
  deleteProperty("progress_num");
}

function onExec() {
  deleteTrigger();
  deleteProperty("add_sheet");
  deleteProperty("cur_page_url");
  deleteProperty("progress_num");
  myFunction(false);
}

function onReExec() {
  myFunction(true);
}

function myFunction(isReExec) {
  // +++++++++++++++++++++++++++++++++++++++++++++++++++++
  // ++ initialize
  // +++++++++++++++++++++++++++++++++++++++++++++++++++++
  var start_time = new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ref_sheet = ss.getSheetByName('シート1');
  // +++++++++++++++++++++++++++++++++++++++++++++++++++++
  // ++ resotre each parameter
  // +++++++++++++++++++++++++++++++++++++++++++++++++++++
  // cur_page_url
  var cur_page_url = !isReExec ?
                     trim(ref_sheet.getRange("A4").getValue()) :
                     getProperty("cur_page_url");
  if (!cur_page_url) { 
    Browser.msgBox("A4セルにurlを入力してください。");
    return;
  }
  Logger.log(cur_page_url);
  
  // progress_num
  var progress_num = !isReExec ?
                     0 :
                     parseInt(getProperty("progress_num")) + 1;
  Logger.log(progress_num);
  
　// add_sheet
  var add_sheet = !isReExec ?
                  ss.insertSheet(Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmmss")) :
                  ss.getSheetByName(getProperty("add_sheet"));
  Logger.log(add_sheet.getSheetName());
  
  // +++++++++++++++++++++++++++++++++++++++++++++++++++++
  // ++ main process start
  // +++++++++++++++++++++++++++++++++++++++++++++++++++++
  while(1) {
    var page_html = UrlFetchApp.fetch(cur_page_url).getContentText();
    Utilities.sleep(1*1000); // after fetched html, must be waited at least 1sec.
   
    var shop_url_list = makeShopUrlList(page_html);
    if (shop_url_list == null) {
　　    Browser.msgBox("店舗リストのurlを取得出来ませんでした")
   　　 return;
    }
    
    for (i = progress_num; i < shop_url_list.length; i++) {
      
      var shop_html = UrlFetchApp.fetch(shop_url_list[i]).getContentText();
      Utilities.sleep(1*1000); // after fetched html, must be waited at least 1sec.

      var shop_name = parseHTML(
        shop_html,
        /<th>店名<\/th>[\s\S]*?<td>([\s\S]*?)<\/td>/
      );
      //Logger.log(shop_name);
      //return;
      
      var reservations_and_inquiries = parseHTML(
        shop_html,
        /<strong class="rstinfo-table__tel-num">([\s\S]*?)<\/strong>/
      );
      //Logger.log(reservations_and_inquiries);
      //return;
      
      var address = deleteTags(
        parseHTML(
          shop_html,
          /<p class="rstinfo-table__address">([\s\S]*?)<\/p>/
        )
      );
      //Logger.log(address);
      //return;

      var business_hours = deleteTags(
        parseHTML(
          shop_html,
          /<th>営業時間<\/th>[\s\S]*?<td>([\s\S]*?)<\/td>/
        )
      );
      //Logger.log(business_hours);
      //return;

      var regular_holiday = deleteTags(
        parseHTML(
          shop_html,
          /<th>定休日<\/th>[\s\S]*?<td>([\s\S]*?)<\/td>/
        )
      );
      //Logger.log(regular_holiday);
      //return;

      var open_date = parseHTML(
        shop_html,
        /<p class="rstinfo-opened-date">([\s\S]*?)<\/p>/
      );
      //Logger.log(open_date);
      //return;

      writeShopData(
        add_sheet,
        shop_url_list[i],
        shop_name,
        reservations_and_inquiries,
        address,
        business_hours,
        regular_holiday,
        open_date
      );
      //return;
      
      // extension gas execute time
      var exec_minutes = parseInt((new Date() - start_time) / (1000 * 60));
      if (exec_minutes >= 4) {
        setProperty("add_sheet", add_sheet.getSheetName());
        setProperty("cur_page_url", cur_page_url);
        setProperty("progress_num", i);
        //setTrigger("myFunction");
        setTrigger("onReExec");
        return;
      }
      //
    }
    var nxt_page_url = parseNextPageUrl(page_html)
    //nxt_page_url = null;
    if (!(nxt_page_url == null)) {
      cur_page_url = nxt_page_url;
      progress_num = 0;
    } else {
      break;
    }
  }
  // +++++++++++++++++++++++++++++++++++++++++++++++++++++
  // terminate
  // +++++++++++++++++++++++++++++++++++++++++++++++++++++
  deleteTrigger();
  deleteProperty("add_sheet");
  deleteProperty("cur_page_url");
  deleteProperty("progress_num");
  Browser.msgBox("終了しました。");
}

// property methods
function setProperty(key, value) {
  deleteProperty(key);
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty(key, value);
}

function getProperty(key) {
  var properties = PropertiesService.getScriptProperties();
  var property = properties.getProperty(key) 
  return !property ? null: property;
}

function deleteProperty(key) {
  var properties = PropertiesService.getScriptProperties();
  properties.deleteProperty(key);
}

// trigger methods    
function setTrigger(func_name) {
  deleteTrigger();
  var dt = new Date();
  dt.setMinutes(dt.getMinutes() + 1); // after 1 minutes, this script will be executed
  ScriptApp.newTrigger(func_name).timeBased().at(dt).create();
}

function deleteTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  if (!triggers) { return; }
  for(var i=0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
} 

// parse methods
function makeShopUrlList(page_html) {
  var shop_url_list = null;
  if (page_html) {
    var tmp_list = page_html.match(/data-detail-url="[\s\S]*?"/g);
    if (tmp_list) {
      if (shop_url_list == null) { shop_url_list = []; }
      tmp_list.forEach(function(tmp, index) {
        var m = tmp.match(/data-detail-url="([\s\S]*?)"/);
        if (m) { shop_url_list[index] = m[1]; }
        //Logger.log("==============================");
        //Logger.log("No." + index);
        //Logger.log(shop_url_list[index]);
        //Logger.log("==============================");
      })
    }
  }
  return shop_url_list
}

function parseHTML(shop_html, pattern) {
  var res = null;
  if (shop_html) {
    var m = shop_html.match(pattern);
    if (m) { res = m[1]; }
    //Logger.log(res);
  }
  return shapeString(trim(res));
}

function parseNextPageUrl(page_html) {
  var res = null;
  if (page_html) {
    var m = page_html.match(/<a href="(.*)" rel="next" class="c-pagination__arrow c-pagination__arrow--next">次の[0-9]+件<\/a>/);
    //Logger.log(m);
    if (m) { res = m[1]; }
    //Logger.log(res);
  }
  return trim(res);
}

function trim(target) {
  if (!target) {
    return target;
  }
  return target.replace(/(^\s+)|(\s+$)/g, "");
}

function deleteTags(target) {
  if (!target) {
    return target;
  }
  return target.replace(/<("[^"]*"|'[^']*'|[^'">])*>/g, "");
}

function shapeString(target) {
  if (!target) {
    return "not found";
  }
  var ary = target.split("\n");
  if (Array.isArray(ary)) {
    for (var i = 0; i <= ary.length-1; i++) {
      ary[i] = trim(ary[i]);
    }
    var appends = "";
    for (var i = 0; i <= ary.length-1; i++) {
      appends = !appends ? (ary[i]): (appends + " " + ary[i]);
    }
    return appends
  } else {
    ary = trim(ary);
    return ary;
  }
}

function writeShopData(
  add_sheet,
  shop_url,
  shop_name,
  reservations_and_inquiries,
  address,
  business_hours,
  regular_holiday,
  open_date
) {
  var used_row = add_sheet.getLastRow();
  //Logger.log(used_row);
  if (!used_row) {
    var base_range = add_sheet.getRange("A1");
    base_range.offset(0,0).setValue("URL");
    base_range.offset(0,1).setValue("店名");
    base_range.offset(0,2).setValue("予約・お問い合わせ");
    base_range.offset(0,3).setValue("住所");
    base_range.offset(0,4).setValue("営業時間");
    base_range.offset(0,5).setValue("定休日");
    base_range.offset(0,6).setValue("オープン日");
  }
  used_row = !used_row ? 1 : used_row;
  var base_range = add_sheet.getRange("A" + (used_row + 1));
  base_range.offset(0,0).setValue(shop_url);
  base_range.offset(0,1).setValue(shop_name);
  base_range.offset(0,2).setValue(reservations_and_inquiries);
  base_range.offset(0,3).setValue(address);
  base_range.offset(0,4).setValue(business_hours);
  base_range.offset(0,5).setValue(regular_holiday);
  base_range.offset(0,6).setValue(open_date);  
}

