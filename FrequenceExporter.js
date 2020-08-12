function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [
        { name: 'Export Frequencies', functionName: 'sdrFreqExport' }
    ];
    spreadsheet.addMenu('SDR Tools', menuItems);
}

id = 1;
categoryid = 1;

function sdrFreqExportSheet(sheet) {
    // This represents ALL the data
    var name = sheet.getName();
    var range = sheet.getDataRange();
    var values = range.getValues();

    var doc = ['<category id="' + categoryid + '" name="' + name + '">'];
    categoryid++;
    for (var i = 2; i < values.length; i++) {
        doc.push('<preset id="' + id + '" ' +
            'name="' + values[i][3] + '" ' +
            'freq="' + rawFreq(values[i][1]) + '" ' +
            'centfreq="' + rawFreq(values[i][1]) + '" ' +
            'offset="0" ' +
            'order="' + (i - 1) + '" ' +
            'filter="' + rawType(values[i][2]) + '" ' +
            'dem="' + rawDim(values[i][2]) + '"/>');
        id++
    }
    doc.push('</category>');
    return doc.join("\n");
}

function sdrFreqExport() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var doc = ['<?xml version="1.0" encoding="UTF-8"?>',
        '<sdr_presets version="1">'];
    for (var i = 0; i < ss.getSheets().length; i++) {
        doc.push(sdrFreqExportSheet(ss.getSheets()[i]));
    }
    doc.push('</sdr_presets>');
    showDoc(doc.join("\n"));
}

function rawFreq(freq, band) {
    return Decimal(freq).mul(1000000);
}

function rawType(type) {
    var dicType = {
        AM: 70000,
        FM: 105000
    }
    return dicType[type];
}

function rawDim(dim) {
    var dicDim = {
        AM: 2,
        FM: 0
    }
    return dicDim[dim];
}

function showDoc(doc) {
    var src = "<pre><![CDATA[" + doc + "]]></pre>";
    var html = HtmlService.createHtmlOutput(src)
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi()
        .showModalDialog(html, 'SDR Export');
}

// Decimal.js (https://github.com/hiroshi-manabe/JSDecimal)
// Minifier with JavaScript Minifier(https://javascript-minifier.com)
!function(){var i=function(r){if(this.constructor!=i)return new i(r);if(!r){this.sig=Array(i.n);for(var o=0;o<i.n;++o)this.sig[o]=0;return this.exp=0,void(this.is_minus=!1)}if(r instanceof i)return this.sig=r.sig.slice(),this.exp=r.exp,void(this.is_minus=r.is_minus);var n="";n=r instanceof String?r:r.toString(),this.is_minus=!1;var t=n.charAt(0);"+"!=t&&"-"!=t||(this.is_minus="-"==t,n=n.substring(1));var e=0,s=n.indexOf("e");if(-1!=s){if(e=-parseInt(n.substring(s+1),10),isNaN(e))throw new i.ArgumentError;n=n.substring(0,s)}var d=n.indexOf(".");-1!=d&&(e+=(n=n.substring(0,d)+n.substring(d+1,n.length)).length-d),n=n.replace(/^0+/,"");var a=i.decimal_digits-(n.length>e?n.length:e),u=0;if(!n.match(/^[0-9]*$/))throw new i.ArgumentError;if(a>0)for(o=0;o<a;++o)n+="0";else if(a<0){var _=n.substring(i.decimal_digits-1,i.decimal_digits),f=n.substring(i.decimal_digits,i.decimal_digits+1),g=parseInt(_),p=parseInt(f);!isNaN(g)&&g%2==1&&!isNaN(p)&&p>=5&&(u=1),n=n.substring(0,i.decimal_digits)}if(e+=a,n.length-e>i.decimal_digits)throw new i.OverflowError;if(n.length<i.decimal_digits){var l="";for(o=0;o<i.decimal_digits-n.length;++o)l+="0";n=l+n}if(this.exp=e,e<-1)throw new i.OverflowError;this.sig=Array();for(o=i.n-1;o>=0;--o){var w=parseFloat(n.substring(o*i.decimal_digits_per_word,(o+1)*i.decimal_digits_per_word))+u;w==i.one_word?w=0:u=0,this.sig.push(w)}if(1==u)throw new i.OverflowError;this.isZero()&&(this.is_minus=!1)};i.pow=function(i,r){for(var o=1,n=0;n<r;++n)o*=i;return o},i.MidpointRounding={AwayFromZero:0,ToEven:1,Floor:2,Ceil:3},i.decimal_digits_per_word=7,i.n=4,i.decimal_digits=i.decimal_digits_per_word*i.n,i.one_word=i.pow(10,i.decimal_digits_per_word),i.two_words=i.pow(10,2*i.decimal_digits_per_word),i.OverflowError=function(){this.message="OverflowError"},i.ZeroDivisionError=function(){this.message="ZeroDivisionError"},i.ArgumentError=function(){this.message="ArgumentError"},i.initConstants=function(){var r="\\d{";r+=i.decimal_digits_per_word,r+="}$",i.regexp_digits=new RegExp(r),i.one_word_zeros="";for(var o=0;o<i.decimal_digits_per_word-1;++o)i.one_word_zeros+="0";i.zeros="";for(o=0;o<i.decimal_digits-1;++o)i.zeros+="0";i.constants_initialized=!0},i.prototype.toString=function(){if(this.isZero())return"0";i.constants_initialized||i.initConstants();for(var r="",o=i.n-1;o>=0;--o)(i.zeros+this.sig[o]).match(i.regexp_digits),r+=RegExp.lastMatch;return this.exp>0&&(r=(r=r.substring(0,i.decimal_digits-this.exp)+"."+r.substring(i.decimal_digits-this.exp)).replace(/\.?0+$/,"")),"."==(r=r.replace(/^0+/,"")).charAt(0)&&(r="0"+r),-1==this.exp&&(r+="0"),(this.is_minus?"-":"")+r},i.prototype.toFloat=function(){return parseFloat(this.toString())},i.fromData=function(r,o,n){var t=new i,e=r.length,s=i.countValidNum(r);if(t.exp=o<s?i.decimal_digits+o-s:i.decimal_digits,t.is_minus=n,t.exp<-1)throw new i.OverflowError;for(var d=Math.floor((o-t.exp+i.decimal_digits)/i.decimal_digits_per_word)-i.n,a=o-t.exp-d*i.decimal_digits_per_word,u=i.pow(10,a),_=i.pow(10,i.decimal_digits_per_word-a),f=0;f<i.n;++f){var g=f+d;g>=0&&g<e&&(t.sig[f]=Math.floor(r[g]/u)),g+1>=0&&g+1<e&&(t.sig[f]+=r[g+1]*_%i.one_word)}if(o>t.exp){var p=!0;for(f=0;f<d+1;++f)if(f<d)f&&r[f-1]&&(p=!1);else{var l=r[f]*_%i.one_word;f>0&&(l+=Math.floor(r[f-1]/u),r[f-1]%u&&(p=!1)),l>i.one_word/2?t.sig[0]+=1:l==i.one_word/2&&(!p||t.sig[0]%2)&&(t.sig[0]+=1)}}return t.isZero()&&(t.is_minus=!1),t},i.Abs=function(r){var o=new i(r);return o.is_minus=!1,o},i.Floor=function(r){return i.RoundInternal(r,0,r.is_minus?i.MidpointRounding.Ceil:i.MidpointRounding.Floor)},i.Ceil=function(r){return i.RoundInternal(r,0,r.is_minus?i.MidpointRounding.Floor:i.MidpointRounding.Ceil)},i.Truncate=function(r){return i.RoundInternal(r,0,i.MidpointRounding.Floor)},i.Round=function(r,o,n){if(n!=i.MidpointRounding.AwayFromZero&&n!=i.MidpointRounding.ToEven)throw new i.ArgumentError;return i.RoundInternal(r,o,n)},i.round=i.Round,i.RoundInternal=function(r,o,n){var t=new i(r),e=r.exp-o;if(e<=0||e>i.decimal_digits||r.isZero())return t;for(var s=Math.floor((e-1)/i.decimal_digits_per_word),d=(e-1)%i.decimal_digits_per_word,a=Math.floor(e/i.decimal_digits_per_word),u=e%i.decimal_digits_per_word,_=i.pow(10,u),f=!1,g=i.pow(10,d+1),p=!0,l=0;l<s;++l)r.sig[l]&&(p=!1,t.sig[l]=0);var w=r.sig[s]%g;if(n==i.MidpointRounding.Floor);else if(n!=i.MidpointRounding.Ceil||0==w&&p){var h=a==i.n||Math.floor(r.sig[a]/_)%2==0;w>g/2?f=!0:w==g/2&&(p&&h&&n==i.MidpointRounding.ToEven||(f=!0))}else f=!0;if(t.sig[s]-=t.sig[s]%g,f)if(a==i.n)t.sig[a-1]=i.one_word/10,--t.exp;else for(t.sig[a]+=_;t.sig[a]==i.one_word;){if(a==i.n-1){t.sig[a]/=10,--t.exp;break}t.sig[a]=0,++a,++t.sig[a]}return t},i.prototype.compare=function(r){var o=r instanceof i?r:new i(r),n=this.is_minus?-1:1;return this.is_minus!=o.is_minus?n:this.absCompare(o)*n},i.prototype.absCompare=function(r){var o=this.isZero(),n=r.isZero();if(o&&n)return 0;if(o!=n)return n-o;var t=r.exp-this.exp;if(t)return t;for(var e=i.n-1;e>=0;--e){var s=this.sig[e]-r.sig[e];if(s)return s}return 0},i.prototype.absAddSub=function(r,o){if(r.isZero())return new i(this);for(var n=Math.floor((r.exp-this.exp)/i.decimal_digits_per_word),t=i.pow(10,(r.exp-this.exp)%i.decimal_digits_per_word),e=o?-1:1,s=o?i.one_word:0,d=o?i.two_words-i.one_word:0,a=i.n+n+1,u=Array(a),_=0;_<a;++_)u[_]=s+d,_>=n&&_-n<i.n&&(u[_]+=this.sig[_-n]*t),_<i.n&&(u[_]+=r.sig[_]*e),s=Math.floor(u[_]/i.one_word),u[_]%=i.one_word;return i.fromData(u,r.exp,this.is_minus)},i.countValidNum=function(r){for(var o=r.length-1;o>=0;--o)for(var n=i.one_word,t=i.decimal_digits_per_word-1;t>=0;--t)if(n/=10,r[o]>=n)return o*i.decimal_digits_per_word+t+1;return 1},i.prototype.validWords=function(){for(var i=this.sig.length-1;i>=0;--i)if(this.sig[i])return i+1;return 1},i.prototype.addSub=function(r,o){var n,t,e,s=r instanceof i?r:new i(r),d=this.absCompare(s);d<0?(n=this,t=s,e=s.is_minus!=o):(n=s,t=this,e=this.is_minus);var a=n.is_minus==t.is_minus==o;if(a&&0==d)return new i(0);var u=t.absAddSub(n,a);return u.is_minus=e,u},i.prototype.add=function(i){return this.addSub(i,!1)},i.prototype.sub=function(i){return this.addSub(i,!0)},i.prototype.isZero=function(){for(var r=0;r<i.n;++r)if(this.sig[r])return!1;return!0},i.prototype.mul=function(r){var o=r instanceof i?r:new i(r);if(this.isZero()||o.isZero())return i(0);for(var n=Array(2*i.n),t=0;t<2*i.n;++t)n[t]=0;for(t=0;t<i.n;++t)for(var e=0;e<i.n;++e){var s=n[t+e]+this.sig[t]*o.sig[e];n[t+e]=s%i.one_word,n[t+e+1]+=Math.floor(s/i.one_word)}var d=this.is_minus!=o.is_minus;return i.fromData(n,this.exp+o.exp,d)},i.prototype.div=function(r){var o=r instanceof i?r:new i(r);if(o.isZero())throw new i.ZeroDivisionError;for(var n=2*i.n,t=Array(n),e=0;e<n;++e)t[e]=0;t=t.concat(this.sig);var s=this.validWords()+n,d=o.validWords(),a=(o.exp,o.sig.slice());1==d&&(a.unshift(0),++d,i.decimal_digits_per_word);var u=i.n+2,_=Array(u);for(e=0;e<u;++e)_[e]=0;var f=0;for(e=0;e<u;++e){var g=s-e-1,p=t[g];e&&(p+=t[g+1]*i.one_word);var l=Math.floor(p/a[d-1]),w=(p-a[d-1]*l)*i.one_word+t[g-1]+(i.two_words-a[d-2]*l);w<i.two_words&&(l-=Math.floor((i.two_words-1-w)/(a[d-1]*i.one_word+a[d-2]))+1);for(var h=i.one_word,c=0;c<d+1;++c){if((m=n-e+c-(i.n-(s-n))+(i.n-d))>=s)break;t[m]+=i.two_words-i.one_word+h,c<d&&(t[m]-=a[c]*l),h=Math.floor(t[m]/i.one_word),t[m]%=i.one_word}if(h<i.one_word)for(l-=1,h=0,c=0;c<d;++c){var m;t[m=s-e-d+c]+=h+a[c],h=Math.floor(t[m]/i.one_word),t[m]%=i.one_word}if(_[u-1-e]=l,(l||f)&&++f,f>i.n)break}for(e=0;e<t.length;++e)if(t[e]){_[0]+=1;break}var v=this.exp-o.exp+(d-(s-n)+i.n+1)*i.decimal_digits_per_word;return i.fromData(_,v,this.is_minus!=o.is_minus)},i.prototype.mod=function(r){var o=r instanceof i?r:new i(r),n=i.Abs(this),t=i.Abs(o),e=n.sub(i.Truncate(n.div(t)).mul(t));return e.is_minus=this.is_minus,e},i.prototype.neg=function(){var r=new i(this);return r.isZero()||(r.is_minus=!r.is_minus),r},"undefined"!=typeof module&&module.exports?module.exports=i:this.Decimal=i}();