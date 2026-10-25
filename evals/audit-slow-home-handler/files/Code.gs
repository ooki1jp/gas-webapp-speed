// Minimal staff-portal backend (fixture for the eval). Deliberately slow in the usual way.
var SHEETS = {
  accounts: { name: "accounts", headers: ["loginId", "name", "empNo"] },
  payslips: {
    name: "payslips",
    headers: ["id", "loginId", "company", "ym", "status"],
  },
  views: { name: "views", headers: ["at", "loginId", "payslipId"] },
  wishes: { name: "wishes", headers: ["loginId", "month", "day"] },
  docs: {
    name: "docs",
    headers: ["id", "loginId", "company", "title", "status"],
  },
};

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var acc = auth_(body.token);
  if (!acc) return json_({ ok: false, error: "auth" });
  if (body.action === "home") return json_(home_(acc));
  return json_({ ok: false, error: "unknown" });
}

function ss_() {
  return SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID"),
  );
}

function readRows_(def) {
  var sh = ss_().getSheetByName(def.name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, def.headers.length).getValues();
  return values.map(function (r) {
    var o = {};
    def.headers.forEach(function (h, i) {
      o[h] = r[i];
    });
    return o;
  });
}

function auth_(token) {
  var id = String(token || "").split(".")[0];
  var accs = readRows_(SHEETS.accounts);
  for (var i = 0; i < accs.length; i++)
    if (accs[i].loginId === id) return accs[i];
  return null;
}

function companiesOf_(acc) {
  var list = [];
  if (acc.empNo) list.push("saito");
  var rows = readRows_(SHEETS.payslips);
  var hasJames = rows.some(function (r) {
    return r.loginId === acc.loginId && r.company === "james";
  });
  if (hasJames) list.push("james");
  return list;
}

function viewedIds_(loginId) {
  var set = {};
  readRows_(SHEETS.views).forEach(function (v) {
    if (v.loginId === loginId) set[v.payslipId] = v.at;
  });
  return set;
}

function listPayslips_(acc, company) {
  var viewed = viewedIds_(acc.loginId);
  return readRows_(SHEETS.payslips)
    .filter(function (r) {
      return (
        r.loginId === acc.loginId &&
        r.company === company &&
        r.status === "public"
      );
    })
    .map(function (r) {
      return { id: r.id, ym: r.ym, viewedAt: viewed[r.id] || null };
    });
}

function docsList_(acc, company) {
  var viewed = viewedIds_(acc.loginId);
  return readRows_(SHEETS.docs)
    .filter(function (r) {
      return (
        r.loginId === acc.loginId &&
        r.company === company &&
        r.status === "public"
      );
    })
    .map(function (r) {
      return {
        id: r.id,
        title: r.title,
        viewedAt: viewed["doc:" + r.id] || null,
      };
    });
}

function home_(acc) {
  var unviewed = [];
  companiesOf_(acc).forEach(function (c) {
    listPayslips_(acc, c).forEach(function (p) {
      if (!p.viewedAt) unviewed.push(p);
    });
  });
  var unviewedDocs = [];
  companiesOf_(acc).forEach(function (c) {
    docsList_(acc, c).forEach(function (d) {
      if (!d.viewedAt) unviewedDocs.push(d);
    });
  });
  return { ok: true, unviewed: unviewed, unviewedDocs: unviewedDocs };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
