/* Manufacturing Proforma — import + editable items table */
(function () {
  const $ = (s, p = document) => p.querySelector(s);
  const $$ = (s, p = document) => [...p.querySelectorAll(s)];
  const money = n => "$" + (Number(n) || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const DUP = JSON.parse(($("#manu-dup-data") || {}).textContent || "null");
  const body = $("#itemsBody");
  if (!body) return;   // ليست صفحة بروفورما التصنيع

  window.switchTab = function (tab) {
    $$(".itab").forEach(b => b.classList.toggle("on", b.dataset.tab === tab));
    $$(".itab-panel").forEach(p => p.classList.add("hidden"));
    $("#panel-" + tab).classList.remove("hidden");
  };

  window.pickCompany = function (el) {
    $$("#companyChips .chip").forEach(c => c.classList.remove("sel"));
    el.classList.add("sel"); $("#company_id").value = el.dataset.id;
  };

  window.addRow = function (it) {
    const tr = document.createElement("tr");
    it = it || {};
    tr.innerHTML = `
      <td class="rn"></td>
      <td><input class="m-model ar-name" value="${(it.model || "").replace(/"/g, '&quot;')}" placeholder="Code 106"></td>
      <td><input class="m-wood" value="${it.wood || "As shown in image"}"></td>
      <td><input class="m-metal" value="${it.metal_color || "As shown in image"}"></td>
      <td><input class="m-handle" value="${it.handle || "As shown in image"}"></td>
      <td><input class="m-acc" value="${it.accessory_color || "As shown in image"}"></td>
      <td><input class="m-right colw" type="number" min="0" value="${it.right || 0}"></td>
      <td><input class="m-left colw" type="number" min="0" value="${it.left || 0}"></td>
      <td class="m-qty num">0</td>
      <td><input class="m-price colp" type="number" min="0" step="0.01" value="${it.unit_price || ""}"></td>
      <td class="m-total num">$0.00</td>
      <td><button type="button" class="x">×</button></td>`;
    body.appendChild(tr);
    $$(".m-right,.m-left,.m-price", tr).forEach(i => i.addEventListener("input", recalc));
    $(".x", tr).addEventListener("click", () => { tr.remove(); recalc(); });
    if (it.qty && !it.right && !it.left) { $(".m-right", tr).value = 0; $(".m-left", tr).value = 0; tr.dataset.fixedQty = it.qty; }
    recalc();
  };

  function fillTable(items) {
    body.innerHTML = "";
    (items || []).forEach(it => addRow(it));
    if (!items || !items.length) addRow();
    recalc();
  }

  window.recalc = function () {
    let qtyAll = 0, totAll = 0;
    $$("#itemsBody tr").forEach((tr, i) => {
      $(".rn", tr).textContent = i + 1;
      const right = +$(".m-right", tr).value || 0, left = +$(".m-left", tr).value || 0;
      const price = +$(".m-price", tr).value || 0;
      const qty = (right || left) ? (right + left) : (+tr.dataset.fixedQty || 0);
      const total = qty * price;
      $(".m-qty", tr).textContent = qty;
      $(".m-total", tr).textContent = money(total);
      qtyAll += qty; totAll += total;
    });
    $("#bQty").textContent = qtyAll;
    $("#bTotal").textContent = money(totAll);
    sync();
  };

  function sync() {
    const items = $$("#itemsBody tr").map(tr => {
      const right = +$(".m-right", tr).value || 0, left = +$(".m-left", tr).value || 0;
      const qty = (right || left) ? (right + left) : (+tr.dataset.fixedQty || 0);
      return {
        model: $(".m-model", tr).value, wood: $(".m-wood", tr).value, metal_color: $(".m-metal", tr).value,
        handle: $(".m-handle", tr).value, accessory_color: $(".m-acc", tr).value,
        right, left, qty, unit_price: +$(".m-price", tr).value || 0
      };
    });
    $("#items_json").value = JSON.stringify(items);
  }

  window.saveCustomer = async function () {
    const p = { name: $("#nc_name").value.trim(), contact: $("#nc_contact").value.trim(), phone: $("#nc_phone").value.trim(),
      country: $("#nc_country").value.trim(), tax_no: $("#nc_tax").value.trim(), currency: $("#nc_currency").value, address: $("#nc_address").value.trim() };
    if (!p.name) { alert("اكتب اسم الزبون"); return; }
    const r = await fetch("/api/customer", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(p) }).then(x => x.json());
    const sel = $("#customer_id"); sel.add(new Option(p.name, r.id, true, true));
    $("#newCust").classList.add("hidden");
  };

  window.uploadExcel = async function () {
    const f = $("#excelFile").files[0];
    if (!f) { alert("اختر ملف أولاً"); return; }
    $("#excelStatus").textContent = "⏳ جارٍ التحليل...";
    const fd = new FormData(); fd.append("file", f);
    const r = await fetch("/manufacturing/parse/excel", { method: "POST", body: fd }).then(x => x.json());
    if (r.error && !r.items.length) { $("#excelStatus").textContent = "⚠️ " + r.error; return; }
    fillTable(r.items);
    $("#excelStatus").textContent = `✅ تم استخراج ${r.items.length} بند${r.error ? " — " + r.error : ""}`;
  };

  window.parsePaste = async function () {
    const text = $("#pasteText").value.trim();
    if (!text) { alert("الصق النص أولاً"); return; }
    $("#pasteStatus").textContent = "⏳ جارٍ التحليل...";
    const r = await fetch("/manufacturing/parse/text", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text }) }).then(x => x.json());
    if (r.error && !r.items.length) { $("#pasteStatus").textContent = "⚠️ " + r.error; return; }
    fillTable(r.items);
    $("#pasteStatus").textContent = `✅ تم استخراج ${r.items.length} بند — راجع كل سطر جيداً`;
  };

  window.uploadImage = async function () {
    const f = $("#imgFile").files[0];
    if (!f) { alert("اختر صورة أولاً"); return; }
    $("#imgStatus").textContent = "⏳ جارٍ استخراج النص...";
    const fd = new FormData(); fd.append("file", f);
    const r = await fetch("/manufacturing/parse/image", { method: "POST", body: fd }).then(x => x.json());
    if (r.text) { $("#ocrText").value = r.text; $("#ocrText").classList.remove("hidden"); $("#reparseBtn").classList.remove("hidden"); }
    if (r.error && !r.items.length) { $("#imgStatus").textContent = "⚠️ " + r.error; return; }
    fillTable(r.items);
    $("#imgStatus").textContent = `✅ تم استخراج ${r.items.length} بند من الصورة — راجعها جيداً`;
  };

  window.reparseOcr = async function () {
    const text = $("#ocrText").value.trim();
    const r = await fetch("/manufacturing/parse/text", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text }) }).then(x => x.json());
    fillTable(r.items);
    $("#imgStatus").textContent = `✅ أُعيد التحليل: ${r.items.length} بند`;
  };

  $("#manuForm").addEventListener("submit", e => {
    sync();
    const items = JSON.parse($("#items_json").value || "[]").filter(i => i.qty > 0);
    if (!items.length) { e.preventDefault(); alert("أضف باباً واحداً على الأقل بكمية أكبر من صفر"); }
  });

  if (DUP && DUP.length) fillTable(DUP); else fillTable([]);
})();
