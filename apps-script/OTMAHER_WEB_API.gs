/**
 * ADISYO CEP RAPORU – OTMAHER SALT OKUNUR WEB API
 * Yalnızca mevcut OTMAHER/Adisyo rapor sekmelerini okur.
 * Adisyo veya Trendyol'a yazmaz; sipariş durumunu değiştirmez; Sheet'e yazmaz.
 */
function doGet(e) {
  try {
    var p=(e&&e.parameter)||{};
    if(String(p.view||'')!=='dashboard') return json_({ok:false,error:'Geçersiz görünüm'});
    var expected=String(PropertiesService.getScriptProperties().getProperty('OTMAHER_WEB_TOKEN')||'');
    var supplied=String(p.token||'');
    if(!expected||!supplied||supplied!==expected) return json_({ok:false,error:'Yetkisiz erişim'});
    return json_(payload_());
  } catch(err) { return json_({ok:false,error:String(err&&err.message||err)}); }
}
function payload_(){
  var ss=SpreadsheetApp.openById('1bcPJFxpg1nPApDAl_Gngp70rGQTOoDnJRtsLLbcYo3s');
  var tz=ss.getSpreadsheetTimeZone()||'Europe/Istanbul';
  var hist=values_(ss,'ADISYO_KAR_GECMIS',17), sys=values_(ss,'ADISYO_SISTEM_DURUMU',4), ham=values_(ss,'ADISYO_HAM',35);
  var history=[];
  for(var i=1;i<hist.length;i++){
    var r=hist[i]; if(!r[0])continue;
    history.push({date:date_(r[0],tz),orders:num_(r[2]),mainProducts:num_(r[3]),revenue:num_(r[4]),productCost:num_(r[7]),platformCut:num_(r[8]),vat:num_(r[9]),preFixedProfit:num_(r[10]),fixedCost:num_(r[11]),temporaryNet:num_(r[12]),definiteNet:r[13]===''||r[13]==null?null:num_(r[13]),accumulatedDefinite:num_(r[14]),issue:String(r[15]||''),status:String(r[16]||'')});
  }
  history.sort(function(a,b){return String(a.date).localeCompare(String(b.date));});
  var overview=history.length?history[history.length-1]:{};
  var sm={}; sys.forEach(function(r){if(r[0])sm[String(r[0]).trim()]=r[1];});
  var system={status:String(sm['Durum']||''),lastAttempt:dt_(sm['Son deneme'],tz),lastSuccess:dt_(sm['Son başarılı senkron'],tz),consecutiveErrors:num_(sm['Ardışık hata']),closedOrders:num_(sm['Kapanmış sipariş']),unmatched:String(sm['Eşleşmeyen']||'Yok')};
  var byId={};
  for(var h=1;h<ham.length;h++){
    var x=ham[h],id=String(x[1]||'').trim(); if(!id)continue;
    var dt=dt_(x[3],tz); if(!dt)continue;
    if(!byId[id]) byId[id]={orderId:id,adisyoOrderNo:String(x[2]||''),integrationOrderId:String(x[19]||''),dateTime:dt,date:date_(x[3],tz),time:time_(x[3],tz),status:String(x[6]||''),orderTotal:num_(x[9]),discount:num_(x[10]),tax:num_(x[11]),paymentMethod:String(x[13]||''),payment:num_(x[14])||num_(x[9]),platform:String(x[18]||x[16]||''),kitchen:kitchen_(x[20]),cancelReason:String(x[21]||''),products:[]};
    var o=byId[id],pn=String(x[25]||'').trim(),q=num_(x[26]);
    if(pn){var pt=(q&&q!==1?q+'× ':'')+pn;if(o.products.indexOf(pt)<0)o.products.push(pt);}
  }
  var orders=Object.keys(byId).map(function(k){return byId[k];}).sort(function(a,b){return String(b.dateTime).localeCompare(String(a.dateTime));}).slice(0,500);
  var cancelled=orders.filter(function(o){return String(o.status).toLocaleLowerCase('tr-TR').indexOf('iptal')>=0;});
  var today=orders.filter(function(o){return o.date===overview.date;});
  return {ok:true,generatedAt:new Date().toISOString(),overview:overview,history:history,system:system,orders:orders,recentOrders:orders.slice(0,30),cancelledOrders:cancelled,todayOrders:today};
}
function values_(ss,n,w){var s=ss.getSheetByName(n);if(!s)return[];var r=s.getLastRow();return r<1?[]:s.getRange(1,1,r,w).getValues();}
function num_(v){if(v===''||v==null)return 0;if(typeof v==='number')return isFinite(v)?v:0;var s=String(v).replace(/₺/g,'').replace(/\s/g,'').replace(/\./g,'').replace(',','.').replace('%','');var n=Number(s);return isFinite(n)?n:0;}
function date_(v,tz){if(!v)return'';if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,tz,'yyyy-MM-dd');var s=String(v).trim(),m=s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);if(m)return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);m=s.match(/^(\d{4}-\d{2}-\d{2})/);return m?m[1]:'';}
function dt_(v,tz){if(!v)return'';if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,tz,"yyyy-MM-dd'T'HH:mm:ssXXX");var d=new Date(v);return isNaN(d.getTime())?String(v):d.toISOString();}
function time_(v,tz){if(!v)return'';if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,tz,'HH:mm');var d=new Date(v);return isNaN(d.getTime())?'':Utilities.formatDate(d,tz,'HH:mm');}
function kitchen_(v){var raw=String(v||''),s=raw.toLocaleLowerCase('tr-TR');if(s.indexOf('acele et')>=0)return'Acele Et Izgara';if(s.indexOf('cızzgara')>=0||s.indexOf('cizzgara')>=0)return'Cızzgara';if(s.indexOf('pilav durağı')>=0||s.indexOf('pilav duragi')>=0||s.indexOf('ekmek arası')>=0||s.indexOf('ekmek arasi')>=0)return'Pilav Durağı';return raw||'Bilinmiyor';}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
