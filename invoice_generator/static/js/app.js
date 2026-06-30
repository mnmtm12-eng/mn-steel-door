/* MN STEEL DOOR — Invoice Generator front-end */
(function () {
  const $ = (s, p = document) => p.querySelector(s);
  const $$ = (s, p = document) => [...p.querySelectorAll(s)];
  const money = n => "$" + (Number(n) || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const modelsEl = $("#models-data");
  if (!modelsEl) return;                       // only on generator page
  const MODELS = JSON.parse(modelsEl.textContent || "[]");
  const DUP = JSON.parse(($("#dup-data") || {}).textContent || "null");
  const byId = {}; MODELS.forEach(m => byId[m.id] = m);

  // ---- model area / width helpers
  function dims(size) {
    const n = (size || "").match(/\d+/g) || [];
    return { w: +n[0] || 100, h: +n[1] || 215 };
  }
  function area(size) { const d = dims(size); return (d.w / 100) * (d.h / 100); }

  // ---------------- rows
  const body = $("#itemsBody");
  let idx = 0;

  function modelOptions(sel) {
    return MODELS.map(m => `<option value="${m.id}" ${+sel === m.id ? "selected" : ""}>${m.code}</option>`).join("");
  }

  window.addRow = function (preset) {
    idx++;
    const tr = document.createElement("tr");
    const mid = preset ? preset.model_id : MODELS[0].id;
    tr.innerHTML = `
      <td class="rn"></td>
      <td><select class="m-model">${modelOptions(mid)}</select></td>
      <td class="m-size mono"></td>
      <td><select class="m-open"><option>Sağ</option><option>Sol</option></select></td>
      <td><input class="m-qty" type="number" min="0" step="1" value="${preset ? preset.qty : 1}"></td>
      <td><input class="m-price" type="number" min="0" step="0.01" value="${preset ? preset.unit_price : ""}"></td>
      <td class="m-total mono">$0.00</td>
      <td class="m-margin secret">—</td>
      <td><button type="button" class="x" title="حذف">×</button></td>`;
    body.appendChild(tr);

    const selModel = $(".m-model", tr), price = $(".m-price", tr), open = $(".m-open", tr);
    if (preset && preset.opening) open.value = preset.opening;

    function fillFromModel() {
      const m = byId[selModel.value];
      $(".m-size", tr).textContent = m.size;
      if (!price.value) price.value = (m.last_price || m.default_price || "") || "";
      recalc();
    }
    selModel.addEventListener("change", fillFromModel);
    $$(".m-qty,.m-price", tr).forEach(i => i.addEventListener("input", recalc));
    $(".x", tr).addEventListener("click", () => { tr.remove(); recalc(); });
    fillFromModel();
  };

  // ---------------- recalc everything
  window.recalc = function () {
    let subtotal = 0, qtyTotal = 0;
    const rows = $$("#itemsBody tr");
    const country = ($("#customer_id").selectedOptions[0] || {}).dataset?.country || "";
    const isIraq = /iraq|عراق|العراق/i.test(country);

    rows.forEach((tr, i) => {
      $(".rn", tr).textContent = i + 1;
      const m = byId[$(".m-model", tr).value];
      const qty = +$(".m-qty", tr).value || 0;
      const price = +$(".m-price", tr).value || 0;
      const total = qty * price;
      subtotal += total; qtyTotal += qty;
      $(".m-total", tr).textContent = money(total);

      // profit guard
      const cost = +m.cost_price || 0;
      const mEl = $(".m-margin", tr);
      if (!cost || !price) { mEl.textContent = "—"; mEl.className = "m-margin secret"; }
      else {
        const pct = ((price - cost) / price) * 100;
        let cls = "g"; if (pct < window.MARGIN.red) cls = "r"; else if (pct < window.MARGIN.yellow) cls = "y";
        mEl.textContent = pct.toFixed(0) + "%";
        mEl.className = "m-margin secret mg-" + cls;
      }
    });

    const advPct = +$("#advance_pct").value || 0;
    const balPct = Math.max(0, 100 - advPct);
    $("#balance_pct").value = balPct; $("#balance_pct_view").value = balPct;
    $("#t_apct").textContent = advPct + "%";
    const adv = subtotal * advPct / 100;
    $("#t_total").textContent = money(subtotal);
    $("#t_adv").textContent = money(adv);
    $("#t_bal").textContent = money(subtotal - adv);

    // commission (secret)
    let comm = 0;
    rows.forEach(tr => {
      const m = byId[$(".m-model", tr).value];
      const qty = +$(".m-qty", tr).value || 0;
      if (isIraq) comm += 3 * area(m.size) * qty;
      else { const w = dims(m.size).w; const amt = w <= 90 ? 5 : (w <= 100 ? 7 : 10); comm += amt * qty; }
    });
    $("#commBox").textContent = money(comm) + (isIraq ? "  ($3/m² · العراق)" : "  (متدرّج)");

    // container calc
    const cap = window.CONTAINER_CAP || 220;
    const cont = Math.ceil(qtyTotal / cap) || 0;
    const rem = qtyTotal % cap;
    let cbox = `${qtyTotal} باب → <b>${cont}</b> كونتينر (السعة ${cap}/كونتينر)`;
    if (qtyTotal > 0 && rem !== 0) {
      const short = cap - rem;
      cbox += `<div class="warn">⚠️ آخر كونتينر ناقص ${short} باب لملئه كاملاً.</div>`;
    }
    $("#contBox").innerHTML = cbox;

    syncItems();
  };

  function syncItems() {
    const items = $$("#itemsBody tr").map(tr => {
      const m = byId[$(".m-model", tr).value];
      return {
        model_id: +m.id, code: m.code, name: m.name, size: m.size,
        opening: $(".m-open", tr).value,
        qty: +$(".m-qty", tr).value || 0,
        unit_price: +$(".m-price", tr).value || 0
      };
    });
    $("#items_json").value = JSON.stringify(items);
  }

  // ---------------- currency
  async function updateFx() {
    const cur = $("#currency").value;
    $("#fx_cur").textContent = cur;
    if (cur === "USD") { $("#t_fx").textContent = "—"; return; }
    try {
      const r = await fetch("/api/rate?to=" + cur).then(x => x.json());
      const subtotal = parseFloat($("#t_total").textContent.replace(/[^0-9.]/g, "")) || 0;
      if (r.rate) $("#t_fx").textContent = (subtotal * r.rate).toLocaleString("en-US", { maximumFractionDigits: 0 }) + " " + cur;
      else $("#t_fx").textContent = "غير متوفر";
    } catch { $("#t_fx").textContent = "—"; }
  }

  // ---------------- NAFEZA (Egypt)
  function checkNafeza() {
    const opt = $("#customer_id").selectedOptions[0];
    const country = (opt ? opt.dataset.country : "") || "";
    const isEg = /egypt|مصر/i.test(country);
    $("#nafeza").classList.toggle("hidden", !isEg);
    $("#acidReq").classList.toggle("hidden", !isEg);
    $("#acid_no").required = isEg;
    // auto currency from customer
    const cur = opt ? opt.dataset.currency : "";
    if (cur) { $("#currency").value = cur; }
    updateFx();
  }

  // ---------------- new customer
  window.toggleNewCust = () => $("#newCust").classList.toggle("hidden");
  window.saveCustomer = async function () {
    const payload = {
      name: $("#nc_name").value.trim(), contact: $("#nc_contact").value.trim(),
      phone: $("#nc_phone").value.trim(), country: $("#nc_country").value.trim(),
      tax_no: $("#nc_tax").value.trim(), currency: $("#nc_currency").value,
      address: $("#nc_address").value.trim()
    };
    if (!payload.name) { alert("اكتب اسم العميل"); return; }
    const r = await fetch("/api/customer", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload)
    }).then(x => x.json());
    const sel = $("#customer_id");
    const o = new Option(payload.name + (payload.country ? " — " + payload.country : ""), r.id, true, true);
    o.dataset.country = payload.country; o.dataset.currency = payload.currency;
    sel.add(o); $("#newCust").classList.add("hidden");
    ["nc_name", "nc_contact", "nc_phone", "nc_country", "nc_tax", "nc_address"].forEach(i => $("#" + i).value = "");
    checkNafeza();
  };

  // ---------------- form guard
  $("#invForm").addEventListener("submit", e => {
    syncItems();
    const items = JSON.parse($("#items_json").value || "[]").filter(i => i.qty > 0);
    if (!items.length) { e.preventDefault(); alert("أضف بنداً واحداً على الأقل بكمية أكبر من صفر"); }
  });

  // ---------------- init
  $("#customer_id").addEventListener("change", () => { checkNafeza(); recalc(); });
  $("#currency").addEventListener("change", updateFx);
  $("#inv_type"); // noop

  if (DUP && DUP.length) DUP.forEach(it => addRow(it));
  else { addRow(); }
  checkNafeza(); recalc();
})();
