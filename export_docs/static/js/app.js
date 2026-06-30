/* AKVA Export Docs — front-end */
(function () {
  const $ = (s, p = document) => p.querySelector(s);
  const $$ = (s, p = document) => [...p.querySelectorAll(s)];
  const money = n => "$" + (Number(n) || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const md = $("#models-data"); if (!md) return;
  const MODELS = JSON.parse(md.textContent || "[]");
  const DUP = JSON.parse(($("#dup-data") || {}).textContent || "null");
  const byId = {}; MODELS.forEach(m => byId[m.id] = m);

  const dBody = $("#doorsBody"), cBody = $("#contBody");

  function opts(sel) { return MODELS.map(m => `<option value="${m.id}" ${+sel === m.id ? "selected" : ""}>${m.name}</option>`).join(""); }

  window.addRow = function (p) {
    const mid = p ? p.model_id : MODELS[0].id; const m = byId[mid] || MODELS[0];
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="rn"></td>
      <td><select class="d-model">${opts(mid)}</select></td>
      <td><input class="d-w colw" value="${p ? p.width : m.width}"></td>
      <td><input class="d-h colw" value="${p ? p.height : m.height}"></td>
      <td><input class="d-b colw" value="${p ? p.beam : m.beam}"></td>
      <td><input class="d-qty colw" type="number" min="0" value="${p ? p.qty : 1}"></td>
      <td><input class="d-price colp" type="number" min="0" step="0.01" value="${p ? p.unit_price : ''}"></td>
      <td class="d-total num">$0.00</td>
      <td><button type="button" class="x">×</button></td>`;
    dBody.appendChild(tr);
    const sm = $(".d-model", tr), pr = $(".d-price", tr);
    sm.addEventListener("change", () => { const mm = byId[sm.value]; $(".d-w", tr).value = mm.width; $(".d-h", tr).value = mm.height; $(".d-b", tr).value = mm.beam; if (!pr.value) pr.value = mm.last_price || mm.default_price || ""; recalc(); });
    $$(".d-qty,.d-price", tr).forEach(i => i.addEventListener("input", recalc));
    $(".x", tr).addEventListener("click", () => { tr.remove(); recalc(); });
    if (!p && !pr.value) pr.value = m.last_price || m.default_price || "";
    recalc();
  };

  window.addCont = function (p) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="rn"></td>
      <td><input class="c-no" value="${p ? (p.no || '') : ''}"></td>
      <td><input class="c-seal" value="${p ? (p.seal || '') : ''}"></td>
      <td><input class="c-content ar-name" value="${p ? (p.content || '') : ''}" placeholder="مثال: 220 باب صلب"></td>
      <td><input class="c-weight colp" type="number" min="0" value="${p ? (p.weight || '') : ''}"></td>
      <td><button type="button" class="x">×</button></td>`;
    cBody.appendChild(tr);
    $$(".c-weight", tr).forEach(i => i.addEventListener("input", recalc));
    $(".x", tr).addEventListener("click", () => { tr.remove(); recalc(); });
    renumber();
  };

  function renumber() { $$("#doorsBody tr").forEach((tr, i) => $(".rn", tr).textContent = i + 1); $$("#contBody tr").forEach((tr, i) => $(".rn", tr).textContent = i + 1); }

  window.recalc = function () {
    let sub = 0, qty = 0;
    $$("#doorsBody tr").forEach(tr => {
      const q = +$(".d-qty", tr).value || 0, p = +$(".d-price", tr).value || 0;
      const t = q * p; sub += t; qty += q; $(".d-total", tr).textContent = money(t);
    });
    let w = 0; $$("#contBody tr").forEach(tr => w += +$(".c-weight", tr).value || 0);
    renumber();
    $("#sumQty").textContent = qty; $("#sumTotal").textContent = money(sub);
    $("#bQty").textContent = qty; $("#bTotal").textContent = money(sub);
    $("#bCont").textContent = $$("#contBody tr").length;
    $("#bWeight").textContent = (w ? w.toLocaleString("en-US") : 0) + " KG";
    sync();
  };

  function sync() {
    $("#items_json").value = JSON.stringify($$("#doorsBody tr").map(tr => {
      const m = byId[$(".d-model", tr).value];
      return { model_id: +m.id, name: m.name, width: $(".d-w", tr).value, height: $(".d-h", tr).value,
        beam: $(".d-b", tr).value, qty: +$(".d-qty", tr).value || 0, unit_price: +$(".d-price", tr).value || 0 };
    }));
    $("#containers_json").value = JSON.stringify($$("#contBody tr").map(tr => ({
      no: $(".c-no", tr).value, seal: $(".c-seal", tr).value, content: $(".c-content", tr).value, weight: +$(".c-weight", tr).value || 0
    })));
  }

  // path preview
  window.buildPath = function () {
    const cust = $("#customer_id").selectedOptions[0]?.textContent.trim() || "[الزبون]";
    const d = $("input[name=date]").value; const ym = d ? d.slice(0, 7) : new Date().toISOString().slice(0, 7);
    const chosen = $$("input[name=docs]:checked").map(c => window.DOC_FOLDERS[c.value].folder);
    const folders = chosen.length ? chosen.join(" · ") : "[نوع الوثيقة]";
    $("#pathPreview").innerHTML = `${window.OUT_BASE}/<b>${cust}</b>/${ym}/<b>${folders}</b>/`;
  };

  function nafeza() {
    const o = $("#customer_id").selectedOptions[0];
    const eg = /egypt|مصر/i.test((o ? o.dataset.country : "") || "");
    $("#nafeza").classList.toggle("hidden", !eg);
    $("#acidReq").classList.toggle("hidden", !eg);
    $("#acid_no").required = eg;
    const cur = o ? o.dataset.currency : ""; if (cur) $("#currency").value = cur;
  }

  window.pickCompany = function (el) {
    $$("#companyChips .chip").forEach(c => c.classList.remove("sel"));
    el.classList.add("sel"); $("#company_id").value = el.dataset.id;
  };

  window.saveCustomer = async function () {
    const p = { name: $("#nc_name").value.trim(), contact: $("#nc_contact").value.trim(), phone: $("#nc_phone").value.trim(),
      country: $("#nc_country").value.trim(), tax_no: $("#nc_tax").value.trim(), currency: $("#nc_currency").value, address: $("#nc_address").value.trim() };
    if (!p.name) { alert("اكتب اسم الزبون"); return; }
    const r = await fetch("/api/customer", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(p) }).then(x => x.json());
    const sel = $("#customer_id"); const o = new Option(p.name, r.id, true, true);
    o.dataset.country = p.country; o.dataset.currency = p.currency; sel.add(o);
    $("#newCust").classList.add("hidden"); nafeza(); buildPath();
  };

  $("#docForm").addEventListener("submit", e => {
    sync();
    if (!JSON.parse($("#items_json").value || "[]").filter(i => i.qty > 0).length) { e.preventDefault(); alert("أضف باباً واحداً على الأقل"); return; }
    if (!$$("input[name=docs]:checked").length) { e.preventDefault(); alert("اختر وثيقة واحدة على الأقل"); }
  });
  $("#customer_id").addEventListener("change", () => { nafeza(); buildPath(); });
  $("input[name=date]").addEventListener("change", buildPath);

  // init
  if (DUP && DUP.items && DUP.items.length) { DUP.items.forEach(it => addRow(it)); (DUP.containers || []).forEach(c => addCont(c)); }
  else { addRow(); addCont(); }
  nafeza(); buildPath(); recalc();
})();
