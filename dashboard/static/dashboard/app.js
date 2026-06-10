(function () {
  "use strict";

  var charts = [];
  var loaded = false;
  var lastReport = null;       // keep the latest report for the warehouse picker
  var whcSelection = "";       // currently selected warehouse in the partner-per-warehouse panel

  var $ = function (id) { return document.getElementById(id); };

  function fmtH(v) {
    if (v == null || isNaN(v)) return "\u2014";
    if (v >= 24) return (v / 24).toFixed(1) + "d";
    return v.toFixed(1) + "h";
  }
  function fmtPct(v) { return v == null || isNaN(v) ? "\u2014" : v.toFixed(1) + "%"; }
  function fmtNum(v, dp) { return v == null || isNaN(v) ? "\u2014" : v.toFixed(dp == null ? 2 : dp); }

  function scoreClass(s) {
    if (s == null) return "na";
    if (s >= 75) return "good";
    if (s >= 50) return "warn";
    return "bad";
  }

  function setHint(msg, kind) {
    var el = $("uploadHint");
    el.textContent = msg || "";
    el.className = "upload-hint" + (kind ? " " + kind : "");
  }

  function destroyCharts() {
    charts.forEach(function (c) { try { c.destroy(); } catch (e) {} });
    charts = [];
  }

  function post(formData) {
    return fetch(window.PROCESS_URL, {
      method: "POST",
      headers: { "X-CSRFToken": window.CSRF_TOKEN || "", "X-Requested-With": "fetch" },
      body: formData
    })
      .then(function (r) {
        // Session expired / not signed in -> bounce to the login page.
        if (r.status === 401 || r.status === 403) {
          window.location.href = window.LOGIN_URL || "/login/";
          throw new Error("Your session has expired. Redirecting to sign in…");
        }
        // Read as text first so a NON-JSON error (e.g. a platform 413
        // "Request Entity Too Large" page, or an HTML 500) gives a readable
        // message instead of an "Unexpected token" JSON-parse crash.
        return r.text().then(function (text) {
          var body;
          try {
            body = text ? JSON.parse(text) : {};
          } catch (e) {
            if (r.status === 413) {
              throw new Error(
                "That file is too large for the server to accept. The host " +
                "is rejecting the upload before it reaches the app (e.g. " +
                "Vercel caps request bodies at 4.5 MB)."
              );
            }
            var snippet = (text || r.statusText || "").trim().slice(0, 160);
            throw new Error("Server error " + r.status + (snippet ? ": " + snippet : ""));
          }
          if (!r.ok) throw new Error(body.error || ("Request failed (" + r.status + ")"));
          return body;
        });
      });
  }

  function currentFilters(fd) {
    fd = fd || new FormData();
    // Each .msel is a multi-select checkbox dropdown; append one field
    // (named by data-name) per checked box. No boxes checked = field omitted
    // = the backend treats it as "all".
    document.querySelectorAll(".msel").forEach(function (m) {
      var name = m.getAttribute("data-name");
      m.querySelectorAll("input[type=checkbox]:checked").forEach(function (cb) {
        fd.append(name, cb.value);
      });
    });
    fd.append("date_from", $("fDateFrom").value);
    fd.append("date_to", $("fDateTo").value);
    return fd;
  }

  function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xlsm|csv|tsv)$/i.test(file.name)) {
      setHint("Please choose a .xlsx or .csv file.", "err");
      return;
    }
    setHint("Reading " + file.name + "\u2026", null);
    var fd = currentFilters();
    fd.append("file", file);
    post(fd)
      .then(function (report) {
        loaded = true;
        setHint("Loaded " + report.summary.total.toLocaleString() + " shipments.", "ok");
        $("uploadZone").hidden = true;
        $("dashboard").hidden = false;
        $("metaStatus").textContent = file.name + " \u00b7 " + report.summary.total.toLocaleString() + " rows";
        render(report);
      })
      .catch(function (err) { setHint(err.message, "err"); });
  }

  function refilter() {
    if (!loaded) return;
    post(currentFilters()).then(render).catch(function (err) {
      $("metaStatus").textContent = "Error: " + err.message;
    });
  }

  // Each dynamic dropdown is filled independently and only the FIRST time it
  // actually receives options. This way a partial first response (e.g. only
  // "zones") never permanently blocks the others from filling on a later one.
  var filled = { zone: false, account: false, warehouse: false };
  var dateBoundsSet = false;

  function populateOptions(filters) {
    filters = filters || {};

    fillMsel("zone", filters.zones, function (z) {
      return { value: z, text: z };
    });
    fillMsel("account", filters.accounts, function (a) {
      return { value: a.value, text: a.label + " (" + (a.n || 0).toLocaleString() + ")" };
    });
    fillMsel("warehouse", filters.warehouses, function (w) {
      return { value: w.value, text: w.label + " (" + (w.n || 0).toLocaleString() + ")" };
    });

    // Bound the date pickers to the data's pickup range (once).
    if (!dateBoundsSet && (filters.date_min || filters.date_max)) {
      var from = $("fDateFrom"), to = $("fDateTo");
      if (filters.date_min) { from.min = filters.date_min; to.min = filters.date_min; }
      if (filters.date_max) { from.max = filters.date_max; to.max = filters.date_max; }
      dateBoundsSet = true;
    }
  }

  // Inject checkbox options into a dynamic .msel dropdown (warehouse/account/zone).
  function fillMsel(name, items, mapFn) {
    if (filled[name]) return;
    items = items || [];
    if (!items.length) return;
    var menu = document.querySelector('.msel[data-name="' + name + '"] .msel-menu');
    if (!menu) return;
    items.forEach(function (it) {
      var o = mapFn(it);
      var lab = document.createElement("label");
      lab.className = "msel-opt";
      var inp = document.createElement("input");
      inp.type = "checkbox"; inp.value = o.value;
      var sp = document.createElement("span");
      sp.textContent = o.text;
      lab.appendChild(inp); lab.appendChild(sp);
      menu.appendChild(lab);
    });
    filled[name] = true;
  }

  // Still used by the warehouse-carrier picker (a single <select>).
  function addOption(sel, value, text) {
    var o = document.createElement("option");
    o.value = value; o.textContent = text;
    sel.appendChild(o);
  }

  // ---- Multi-select dropdown (.msel) behaviour --------------------------

  // Update a dropdown's button label: "All", the single picked option, or
  // "N selected".
  function mselText(m) {
    var checked = m.querySelectorAll("input[type=checkbox]:checked");
    var txt = m.querySelector(".msel-text");
    if (!checked.length) {
      txt.textContent = m.getAttribute("data-all") || "All";
      m.classList.remove("has-sel");
      return;
    }
    m.classList.add("has-sel");
    if (checked.length === 1) {
      var sp = checked[0].parentNode.querySelector("span");
      txt.textContent = sp ? sp.textContent : checked[0].value;
    } else {
      txt.textContent = checked.length + " selected";
    }
  }

  function closeAllMsel(except) {
    document.querySelectorAll(".msel.open").forEach(function (m) {
      if (m !== except) {
        m.classList.remove("open");
        var menu = m.querySelector(".msel-menu");
        if (menu) menu.hidden = true;
      }
    });
  }

  function render(report) {
    lastReport = report;
    populateOptions(report.filters);
    renderStats(report.summary);
    renderFormula(report.weights);
    renderTable(report.carriers);
    renderStatusMatrix(report);
    renderProducts(report);
    renderWarehouses(report);
    renderCities(report);
    renderDestinations(report);
    renderLanes(report);
    renderWarehouseCarrierPanel(report);
    destroyCharts();
    renderCharts(report);
    renderWarehouseChart(report.warehouses);
    renderDestChart(report.destinations);
    renderWhcChart();
  }

  function renderStats(s) {
    var cards = [
      ["Shipments", s.total.toLocaleString(), ""],
      ["Delivery success", fmtPct(s.success_rate), ""],
      ["Avg pickup\u2192OFD1", fmtH(s.avg_p2o), ""],
      ["Avg pickup\u2192delivery", fmtH(s.avg_p2d), ""],
      ["Carriers", String(s.carriers), ""],
      ["Warehouses", String(s.warehouses), ""]
    ];
    $("statRow").innerHTML = cards.map(function (c) {
      return '<div class="stat"><div class="stat-label">' + c[0] +
        '</div><div class="stat-value">' + c[1] + '</div></div>';
    }).join("");
  }

  function renderFormula(w) {
    $("scoreFormula").textContent =
      "Weighted 0\u2013100 \u00b7 " + w.p2o + "% P\u2192OFD1 \u00b7 " + w.p2d +
      "% P\u2192deliv \u00b7 " + w.succ + "% success \u00b7 " + w.fa +
      "% 1st-attempt \u00b7 " + w.att + "% avg attempts";
  }

  // Shared KPI cells used by carrier / warehouse / city / whc tables.
  function kpiCells(a) {
    return '<td class="num">' + a.n.toLocaleString() + "</td>" +
      '<td class="num">' + fmtH(a.p2o) + "</td>" +
      '<td class="num">' + fmtH(a.p2d) + "</td>" +
      '<td class="num">' + fmtPct(a.success_rate) + "</td>" +
      '<td class="num">' + fmtPct(a.first_attempt_rate) + "</td>" +
      '<td class="num">' + fmtNum(a.avg_attempts, 2) + "</td>";
  }

  function scoreCellHtml(score, minN) {
    var cls = scoreClass(score);
    if (score == null) {
      var title = minN ? ' title="Below ' + minN + '-shipment threshold"' : "";
      return '<span class="g-na"' + title + ">\u2014</span>";
    }
    var bar = Math.round(score);
    return '<div class="score-cell"><span class="score-val g-' + cls + '">' +
      score.toFixed(0) + '</span><span class="score-bar"><span class="b-' + cls +
      '" style="width:' + bar + '%"></span></span></div>';
  }

  function renderTable(carriers) {
    var head = "<thead><tr><th></th><th>Carrier</th><th>Score</th><th>Shipments</th>" +
      "<th>P\u2192OFD1</th><th>P\u2192Deliv</th><th>Success</th><th>1st-attempt</th><th>Avg att.</th></tr></thead>";
    var rows = carriers.map(function (a, i) {
      return "<tr>" +
        '<td class="rank">' + (i + 1) + "</td>" +
        '<td class="carrier-name">' + esc(a.carrier) + "</td>" +
        "<td>" + scoreCellHtml(a.score) + "</td>" +
        kpiCells(a) +
        "</tr>";
    }).join("");
    $("scoreTable").innerHTML = head + "<tbody>" + rows + "</tbody>";
  }

  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }

  function renderWarehouses(report) {
    var wh = report.warehouses;
    var minN = report.warehouse_min_n;
    var total = report.warehouse_total != null ? report.warehouse_total : wh.length;
    var topN = report.warehouse_top_n;
    var scored = wh.filter(function (w) { return w.score != null; }).length;
    var lead = (topN != null && total > wh.length)
      ? ("top " + wh.length + " of " + total.toLocaleString() + " pickup pincodes")
      : (total.toLocaleString() + " pickup pincodes");
    $("whSub").textContent =
      lead + " \u00b7 " + scored + " scored (\u2265" + minN +
      " shipments) \u00b7 ranked by volume";

    var head = "<thead><tr><th></th><th>Warehouse</th><th>Score</th><th>Shipments</th>" +
      "<th>P\u2192OFD1</th><th>P\u2192Deliv</th><th>Success</th><th>1st-attempt</th><th>Avg att.</th></tr></thead>";
    var rows = wh.map(function (w, i) {
      return "<tr>" +
        '<td class="rank">' + (i + 1) + "</td>" +
        '<td class="carrier-name">' + esc(w.warehouse || w.pickup_pin) + "</td>" +
        "<td>" + scoreCellHtml(w.score, minN) + "</td>" +
        kpiCells(w) +
        "</tr>";
    }).join("");
    $("whTable").innerHTML = head + "<tbody>" + rows + "</tbody>";
  }

  function renderCities(report) {
    var cities = report.cities || [];
    var minN = report.warehouse_min_n;
    var scored = cities.filter(function (c) { return c.score != null; }).length;
    $("citySub").textContent =
      cities.length + " cities/regions \u00b7 " + scored + " scored (\u2265" + minN +
      " shipments) \u00b7 ranked by volume";

    var head = "<thead><tr><th></th><th>City / region</th><th>Score</th><th>Shipments</th>" +
      "<th>P\u2192OFD1</th><th>P\u2192Deliv</th><th>Success</th><th>1st-attempt</th><th>Avg att.</th></tr></thead>";
    var rows = cities.map(function (c, i) {
      return "<tr>" +
        '<td class="rank">' + (i + 1) + "</td>" +
        '<td class="carrier-name">' + esc(c.city) + "</td>" +
        "<td>" + scoreCellHtml(c.score, minN) + "</td>" +
        kpiCells(c) +
        "</tr>";
    }).join("");
    $("cityTable").innerHTML = head + "<tbody>" + rows + "</tbody>";
  }

  // Cap how many search matches we render at once (keeps the DOM light when a
  // short query matches thousands of rows).
  var SEARCH_CAP = 250;

  // Shared renderer for the destination / lane breakdown tables. Both receive
  // the FULL list from the backend: with no query we show the top-N by volume;
  // with a query we filter the whole list (case-insensitive substring) and
  // keep each row's true volume rank.
  function renderRollup(opts) {
    var all = opts.items || [];
    var minN = opts.minN;
    var topN = opts.topN;
    var total = opts.total != null ? opts.total : all.length;
    var q = (($(opts.searchId).value) || "").trim().toLowerCase();

    // True volume rank (list is already sorted by volume desc).
    var rankOf = {};
    all.forEach(function (d, i) { rankOf[d[opts.labelKey]] = i + 1; });

    var matches = q
      ? all.filter(function (d) {
          return String(d[opts.labelKey] || "").toLowerCase().indexOf(q) !== -1;
        })
      : all.slice(0, topN || all.length);
    var shown = matches.slice(0, SEARCH_CAP);

    if (q) {
      $(opts.subId).textContent =
        matches.length.toLocaleString() + " match" + (matches.length === 1 ? "" : "es") +
        " of " + total.toLocaleString() + " " + opts.noun +
        (matches.length > shown.length ? " · showing first " + shown.length : "");
    } else {
      var scored = all.filter(function (d) { return d.score != null; }).length;
      var lead = (topN != null && total > shown.length)
        ? ("top " + shown.length + " of " + total.toLocaleString() + " " + opts.noun)
        : (total.toLocaleString() + " " + opts.noun);
      $(opts.subId).textContent =
        lead + " · " + scored + " scored (≥" + minN + " shipments) · " +
        opts.tail + " · ranked by volume";
    }

    var head = "<thead><tr><th></th><th>" + opts.colLabel + "</th><th>Score</th><th>Shipments</th>" +
      "<th>P→OFD1</th><th>P→Deliv</th><th>Success</th><th>1st-attempt</th><th>Avg att.</th></tr></thead>";

    if (!shown.length) {
      $(opts.tableId).innerHTML = head +
        '<tbody><tr><td class="empty" colspan="9">No ' + opts.noun.replace(/s$/, "") +
        ' matches “' + esc(q) + '”.</td></tr></tbody>';
      return;
    }

    var rows = shown.map(function (d) {
      return "<tr>" +
        '<td class="rank">' + (rankOf[d[opts.labelKey]] || "") + "</td>" +
        '<td class="carrier-name">' + esc(d[opts.labelKey]) + "</td>" +
        "<td>" + scoreCellHtml(d.score, minN) + "</td>" +
        kpiCells(d) +
        "</tr>";
    }).join("");
    $(opts.tableId).innerHTML = head + "<tbody>" + rows + "</tbody>";
  }

  function renderDestinations(report) {
    renderRollup({
      items: report.destinations, total: report.destination_total,
      topN: report.destination_top_n, minN: report.warehouse_min_n,
      labelKey: "drop_city", noun: "destinations", colLabel: "Destination",
      tail: "by drop city", subId: "destSub", tableId: "destTable",
      searchId: "destSearch",
    });
  }

  function renderLanes(report) {
    renderRollup({
      items: report.lanes, total: report.lane_total,
      topN: report.lane_top_n, minN: report.lane_min_n,
      labelKey: "lane", noun: "lanes", colLabel: "Lane",
      tail: "pickup city → drop city", subId: "laneSub", tableId: "laneTable",
      searchId: "laneSearch",
    });
  }

  // ---- Outcome & pendency matrices --------------------------------------

  // Color a percentage cell: green for the "good" outcome (Delivered),
  // amber/red ramp for problem outcomes by magnitude.
  function pctCell(value, kind) {
    if (value == null) return '<td class="num pct">\u2014</td>';
    var cls = "";
    if (kind === "good") {
      cls = value >= 90 ? "g-good" : value >= 75 ? "g-warn" : "g-bad";
    } else {
      // problem outcome: higher = worse
      cls = value >= 15 ? "g-bad" : value >= 5 ? "g-warn" : "";
    }
    return '<td class="num pct ' + cls + '">' + value.toFixed(1) + "%</td>";
  }

  function renderStatusMatrix(report) {
    var matrix = report.status_matrix || [];
    var outcomes = report.outcomes || ["Delivered", "FWD Pendency", "RTO", "Cancelled"];
    var totalN = matrix.reduce(function (a, r) { return a + r.n; }, 0);
    $("statusSub").textContent =
      matrix.length + " carrier accounts \u00b7 " + totalN.toLocaleString() +
      " shipments \u00b7 row % of each account's own total";

    var head = "<thead><tr><th></th><th>Carrier account</th><th>Shipments</th>";
    outcomes.forEach(function (o) { head += "<th>" + esc(o) + "</th>"; });
    head += "</tr></thead>";

    var rows = matrix.map(function (r, i) {
      var cells = "";
      outcomes.forEach(function (o) {
        var kind = (o === "Delivered") ? "good" : "bad";
        cells += pctCell(r.pct[o], kind);
      });
      return "<tr>" +
        '<td class="rank">' + (i + 1) + "</td>" +
        '<td class="carrier-name">' + esc(r.account) + "</td>" +
        '<td class="num">' + r.n.toLocaleString() + "</td>" +
        cells + "</tr>";
    }).join("");
    $("statusTable").innerHTML = head + "<tbody>" + rows + "</tbody>";
  }

  // ---- Product breakdown ------------------------------------------------

  function renderProducts(report) {
    var products = report.products || [];
    var total = products.reduce(function (a, c) { return a + c.n; }, 0);
    var grid = $("productGrid");

    if (!report.has_products) {
      $("productSub").textContent =
        "This file has no 'Item Names' column \u2014 product breakdown unavailable.";
      grid.innerHTML =
        '<div class="product-empty">Upload an export that includes the Item Names column to see product categories.</div>';
      return;
    }
    $("productSub").textContent =
      products.length + " categories \u00b7 " + total.toLocaleString() +
      " shipments \u00b7 grouped from product names";

    // Color cycle for category accents.
    var ACCENTS = ["#b4451f", "#3d6b2e", "#9a7b1f", "#4a463e", "#6b8aa0", "#7a2d12"];

    grid.innerHTML = products.map(function (c, i) {
      var accent = ACCENTS[i % ACCENTS.length];
      var pct = total ? (c.n / total * 100) : 0;
      var maxSub = c.subs.reduce(function (m, s) { return Math.max(m, s.n); }, 0) || 1;
      var maxAcct = (c.accounts || []).reduce(function (m, a) { return Math.max(m, a.n); }, 0) || 1;

      // Subcategory rows with proportional bars.
      var subHtml = c.subs.map(function (s) {
        var w = Math.round(s.n / maxSub * 100);
        var sp = c.n ? (s.n / c.n * 100) : 0;
        return '<div class="pc-row">' +
          '<span class="pc-row-label" title="' + esc(s.subcategory) + '">' + esc(s.subcategory) + "</span>" +
          '<span class="pc-bar"><span class="pc-bar-fill" style="width:' + w + '%;background:' + accent + '"></span></span>' +
          '<span class="pc-row-val">' + s.n.toLocaleString() + '<i>' + sp.toFixed(0) + "%</i></span>" +
          "</div>";
      }).join("");

      // Top carrier accounts (cap at 6, roll the rest into "others").
      var accts = (c.accounts || []).slice();
      var shown = accts.slice(0, 6);
      var rest = accts.slice(6);
      if (rest.length) {
        var restN = rest.reduce(function (a, x) { return a + x.n; }, 0);
        shown.push({ account: "+" + rest.length + " others", n: restN, _muted: true });
      }
      var acctHtml = shown.map(function (a) {
        var w = Math.round(a.n / maxAcct * 100);
        var ap = c.n ? (a.n / c.n * 100) : 0;
        return '<div class="pc-row' + (a._muted ? " pc-muted" : "") + '">' +
          '<span class="pc-row-label" title="' + esc(a.account) + '">' + esc(a.account) + "</span>" +
          '<span class="pc-bar"><span class="pc-bar-fill" style="width:' + w + '%;background:' + accent + '99"></span></span>' +
          '<span class="pc-row-val">' + a.n.toLocaleString() + '<i>' + ap.toFixed(0) + "%</i></span>" +
          "</div>";
      }).join("");

      return '<div class="pc-card">' +
          '<div class="pc-head" style="--accent:' + accent + '">' +
            '<span class="pc-rank">' + (i + 1) + "</span>" +
            '<span class="pc-name">' + esc(c.category) + "</span>" +
            '<span class="pc-count">' + c.n.toLocaleString() +
              ' <em>' + pct.toFixed(1) + "%</em></span>" +
          "</div>" +
          '<div class="pc-share"><span style="width:' + pct.toFixed(1) + '%;background:' + accent + '"></span></div>' +
          '<div class="pc-section">' +
            '<div class="pc-section-title">Subcategories</div>' + subHtml +
          "</div>" +
          '<div class="pc-section">' +
            '<div class="pc-section-title">Shipped by</div>' + acctHtml +
          "</div>" +
        "</div>";
    }).join("");
  }

  // Group the flat warehouse_carrier list into { pin: { label, rows[] } }.
  function groupWhc(report) {
    var groups = {};
    var order = [];
    (report.warehouse_carrier || []).forEach(function (r) {
      var pin = r.pickup_pin;
      if (!groups[pin]) {
        groups[pin] = { pin: pin, label: r.warehouse || pin, total: r.wh_total, rows: [] };
        order.push(pin);
      }
      groups[pin].rows.push(r);
    });
    // Order warehouses by total volume desc.
    order.sort(function (a, b) { return (groups[b].total || 0) - (groups[a].total || 0); });
    return { groups: groups, order: order };
  }

  function renderWarehouseCarrierPanel(report) {
    var g = groupWhc(report);
    var pick = $("whcPick");

    // (Re)build the warehouse picker each render so it tracks the active filter set.
    pick.innerHTML = "";
    g.order.forEach(function (pin) {
      var grp = g.groups[pin];
      addOption(pick, pin, grp.label + " (" + (grp.total || 0).toLocaleString() + ")");
    });

    // Preserve the previous selection if it still exists, else default to the
    // highest-volume warehouse.
    if (!whcSelection || !g.groups[whcSelection]) {
      whcSelection = g.order.length ? g.order[0] : "";
    }
    if (whcSelection) pick.value = whcSelection;

    var minN = report.wh_carrier_min_n;
    var nWh = g.order.length;
    $("whcSub").textContent =
      nWh + " warehouse" + (nWh === 1 ? "" : "s") +
      " \u00b7 carriers scored against each other within each (\u2265" + minN + " shipments per cell)";

    renderWhcTable(report);
  }

  function renderWhcTable(report) {
    var g = groupWhc(report);
    var grp = g.groups[whcSelection];
    var minN = report.wh_carrier_min_n;
    if (!grp) {
      $("whcTable").innerHTML =
        '<tbody><tr><td class="empty">No warehouse data for the current filters.</td></tr></tbody>';
      return;
    }

    var head = "<thead><tr><th></th><th>Carrier</th><th>Score</th><th>Shipments</th><th>Share</th>" +
      "<th>P\u2192OFD1</th><th>P\u2192Deliv</th><th>Success</th><th>1st-attempt</th><th>Avg att.</th></tr></thead>";
    var rows = grp.rows.map(function (a, i) {
      return "<tr>" +
        '<td class="rank">' + (i + 1) + "</td>" +
        '<td class="carrier-name">' + esc(a.carrier) + "</td>" +
        "<td>" + scoreCellHtml(a.score, minN) + "</td>" +
        '<td class="num">' + a.n.toLocaleString() + "</td>" +
        '<td class="num">' + fmtPct(a.wh_share) + "</td>" +
        '<td class="num">' + fmtH(a.p2o) + "</td>" +
        '<td class="num">' + fmtH(a.p2d) + "</td>" +
        '<td class="num">' + fmtPct(a.success_rate) + "</td>" +
        '<td class="num">' + fmtPct(a.first_attempt_rate) + "</td>" +
        '<td class="num">' + fmtNum(a.avg_attempts, 2) + "</td>" +
        "</tr>";
    }).join("");
    $("whcTable").innerHTML = head + "<tbody>" + rows + "</tbody>";
  }

  function renderWhcChart() {
    if (!lastReport) return;
    if (whcChart) { try { whcChart.destroy(); } catch (e) {} whcChart = null; }
    var g = groupWhc(lastReport);
    var grp = g.groups[whcSelection];
    if (!grp) return;
    var labels = grp.rows.map(function (a) { return a.carrier; });
    // Bars = score where available, else 0; colour by score band.
    whcChart = new Chart($("cWhc"), {
      type: "bar",
      data: {
        labels: labels,
        datasets: [{
          label: "Score",
          data: grp.rows.map(function (a) { return a.score == null ? 0 : a.score; }),
          backgroundColor: grp.rows.map(function (a) { return scoreColor(a.score); }),
          borderRadius: 3, barThickness: "flex", maxBarThickness: 24
        }]
      },
      options: hbarOpts()
    });
  }

  function scoreColor(s) {
    if (s == null) return "#b4b2a9";
    if (s >= 75) return "#3d6b2e";
    if (s >= 50) return "#9a7b1f";
    return "#b4451f";
  }

  // -----------------------------------------------------------------------

  function renderWarehouseChart(wh) {
    // Top warehouses by volume, P->delivery TAT as the bar.
    var top = wh.slice(0, 12);
    var labels = top.map(function (w) { return w.warehouse || w.pickup_pin; });
    charts.push(new Chart($("cWh"), {
      type: "bar",
      data: {
        labels: labels,
        datasets: [{
          data: top.map(function (w) { return w.p2d == null ? 0 : w.p2d; }),
          backgroundColor: top.map(function (w) {
            return w.score == null ? "#b4b2a9" : "#9a7b1f";
          }),
          borderRadius: 3, barThickness: "flex", maxBarThickness: 22
        }]
      },
      options: hbarOpts()
    }));
  }

  function renderDestChart(dest) {
    // Top destinations by volume, P->delivery TAT as the bar (blue accent to
    // distinguish from the warehouse/pickup-side chart).
    if (!dest) return;
    var top = dest.slice(0, 12);
    var labels = top.map(function (d) { return d.drop_city; });
    charts.push(new Chart($("cDest"), {
      type: "bar",
      data: {
        labels: labels,
        datasets: [{
          data: top.map(function (d) { return d.p2d == null ? 0 : d.p2d; }),
          backgroundColor: top.map(function (d) {
            return d.score == null ? "#b4b2a9" : "#6b8aa0";
          }),
          borderRadius: 3, barThickness: "flex", maxBarThickness: 22
        }]
      },
      options: hbarOpts()
    }));
  }

  var INK = "#1c1a17", FAINT = "#8a8377", LINE = "rgba(28,26,23,.08)";
  var PALETTE = ["#b4451f", "#3d6b2e", "#9a7b1f", "#4a463e", "#7a2d12", "#6b8aa0"];
  var whcChart = null;

  function hbarOpts() {
    return {
      indexAxis: "y", responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { color: LINE }, ticks: { color: FAINT, font: { family: "Spline Sans Mono" } } },
        y: { grid: { display: false }, ticks: { color: INK, font: { family: "Archivo", size: 12 } } }
      }
    };
  }
  function doughnutOpts() {
    return {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: "bottom", labels: { color: FAINT, boxWidth: 10, font: { family: "Spline Sans Mono", size: 11 } } } },
      cutout: "58%"
    };
  }

  function renderCharts(report) {
    var c = report.carriers;
    var labels = c.map(function (a) { return a.carrier; });

    charts.push(new Chart($("cScore"), {
      type: "bar",
      data: { labels: labels, datasets: [{ data: c.map(function (a) { return a.score == null ? 0 : a.score; }), backgroundColor: "#b4451f", borderRadius: 3, barThickness: "flex", maxBarThickness: 22 }] },
      options: hbarOpts()
    }));

    charts.push(new Chart($("cTat"), {
      type: "bar",
      data: { labels: labels, datasets: [{ data: c.map(function (a) { return a.p2d == null ? 0 : a.p2d; }), backgroundColor: "#3d6b2e", borderRadius: 3, barThickness: "flex", maxBarThickness: 22 }] },
      options: hbarOpts()
    }));

    mixChart("cWeight", report.mix.weight);
    mixChart("cPay", report.mix.payment);
    mixChart("cDtype", report.mix.delivery_type);
  }

  function mixChart(id, obj) {
    var labels = Object.keys(obj);
    var data = labels.map(function (k) { return obj[k]; });
    charts.push(new Chart($(id), {
      type: "doughnut",
      data: { labels: labels, datasets: [{ data: data, backgroundColor: PALETTE, borderColor: "#faf8f2", borderWidth: 2 }] },
      options: doughnutOpts()
    }));
  }

  // Wiring
  $("fileInput").addEventListener("change", function (e) { handleFile(e.target.files[0]); });
  ["fDateFrom", "fDateTo"].forEach(function (id) {
    $(id).addEventListener("change", refilter);
  });

  // Multi-select dropdowns: open/close on the toggle, close others, and close
  // any open menu when clicking outside. Delegated so dynamically-injected
  // options work too.
  document.addEventListener("click", function (e) {
    var toggle = e.target.closest(".msel-toggle");
    if (toggle) {
      var m = toggle.parentNode;
      var willOpen = !m.classList.contains("open");
      closeAllMsel(m);
      m.classList.toggle("open", willOpen);
      m.querySelector(".msel-menu").hidden = !willOpen;
      return;
    }
    if (e.target.closest(".msel-menu")) return;  // clicks inside stay open
    closeAllMsel(null);
  });

  // Any checkbox toggle inside a dropdown: refresh its label and refilter.
  document.addEventListener("change", function (e) {
    var cb = e.target;
    if (cb && cb.matches && cb.matches(".msel input[type=checkbox]")) {
      mselText(cb.closest(".msel"));
      refilter();
    }
  });

  // Initialise each dropdown's button label (handles the default-checked
  // Forward delivery type).
  document.querySelectorAll(".msel").forEach(function (m) { mselText(m); });

  // Per-panel search: filter the (full) destination / lane lists client-side,
  // re-rendering from the cached report so no server round-trip is needed.
  $("destSearch").addEventListener("input", function () {
    if (lastReport) renderDestinations(lastReport);
  });
  $("laneSearch").addEventListener("input", function () {
    if (lastReport) renderLanes(lastReport);
  });

  $("whcPick").addEventListener("change", function (e) {
    whcSelection = e.target.value;
    if (lastReport) { renderWhcTable(lastReport); renderWhcChart(); }
  });
  $("reuploadBtn").addEventListener("click", function () {
    $("dashboard").hidden = true;
    $("uploadZone").hidden = false;
    $("metaStatus").textContent = "No data loaded";
    setHint("", null);
    $("fileInput").value = "";
  });

  var drop = $("dropArea");
  ["dragenter", "dragover"].forEach(function (ev) {
    drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add("drag"); });
  });
  ["dragleave", "drop"].forEach(function (ev) {
    drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove("drag"); });
  });
  drop.addEventListener("drop", function (e) {
    if (e.dataTransfer.files && e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });
})();