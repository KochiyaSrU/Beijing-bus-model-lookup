"""单文件网页版导出。

这个模块只依赖标准库和 Pillow，不认识 tkinter，也不认识 App：调用方把数据整理成
一个普通字典（payload），这里负责把它渲染成一个可以双击打开、也能在手机上看的
HTML 文件（数据与脚本全部内联，不联网、不加载 CDN）。

payload 结构::

    {
        "title": "北京公交车型查询器",
        "generated": "2026-09-30 14:00:00",
        "ranges": [[20000, 20212, "BJ6123C7C4D"], ...],          # 公交号段
        "tags": {"BJ6123C7C4D": ["北汽福田", "柴油", "12 米级"]},  # 厂商/燃料/车长
        "aliases": {"BJ6123C7C4D": ["大福后"]},
        "profiles": {"BJ6123C7C4D": {"intro": "...", "spec": "...", "image": "data:image/jpeg;base64,..."}},
        "metro": {"1号线": [[1001, 1050, "SFM04"], ...]},
        "notices": ["地铁数据与种子文件有 4 处不一致"],
    }
"""

from __future__ import annotations

import base64
import io
import json
from pathlib import Path

MAX_THUMBNAIL = (360, 260)
THUMBNAIL_QUALITY = 72

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
  :root {
    --bg: #f0f3f8; --surface: #fff; --border: #dce3ec; --stripe: #f7fafd;
    --text: #16212e; --muted: #6b7b8f; --primary: #1668c4; --primary-soft: #e8f1fc;
    --accent: #c92936; --ok: #1e7f4f;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 16px; background: var(--bg); color: var(--text);
    font: 15px/1.6 "Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", system-ui, sans-serif;
  }
  .wrap { max-width: 900px; margin: 0 auto; }
  header { background: #123a66; color: #fff; border-radius: 10px 10px 0 0; padding: 14px 18px; }
  header h1 { margin: 0; font-size: 19px; }
  header p { margin: 4px 0 0; font-size: 12px; color: #a9c6e4; }
  .accent { height: 3px; background: #2e7bd6; border-radius: 0 0 10px 10px; margin-bottom: 14px; }
  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 14px 16px; margin-bottom: 12px; }
  h2 { font-size: 13px; color: var(--muted); margin: 0 0 10px; letter-spacing: .5px; }
  input, select {
    font: inherit; padding: 9px 10px; border: 1px solid #c9d4e2; border-radius: 7px;
    background: #fff; color: var(--text); outline: none;
  }
  input:focus, select:focus { border-color: var(--primary); box-shadow: 0 0 0 2px var(--primary-soft); }
  button {
    font: inherit; font-weight: 600; padding: 9px 16px; border: 1px solid var(--border);
    border-radius: 7px; background: #fff; color: var(--text); cursor: pointer;
  }
  button.primary { background: var(--primary); border-color: var(--primary); color: #fff; }
  .row { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
  .grow { flex: 1 1 160px; }
  #number { font-size: 22px; font-weight: 700; text-align: center; letter-spacing: 2px; }
  .result { text-align: center; padding: 10px 0 2px; }
  .result .model { font-size: 26px; font-weight: 700; color: var(--primary); word-break: break-all; }
  .result .hint { color: var(--muted); font-size: 13px; margin-top: 4px; }
  .result .missing { color: var(--muted); font-size: 20px; font-weight: 700; }
  .result .bad { color: var(--accent); font-size: 18px; font-weight: 700; }
  .chips { display: flex; gap: 6px; justify-content: center; flex-wrap: wrap; margin-top: 8px; }
  .chip { background: var(--primary-soft); color: var(--primary); border-radius: 20px; padding: 2px 10px; font-size: 12px; font-weight: 600; }
  table { width: 100%; border-collapse: collapse; font-size: 14px; }
  th { text-align: left; color: var(--muted); font-size: 12px; padding: 8px 6px; border-bottom: 1px solid var(--border); cursor: pointer; white-space: nowrap; }
  td { padding: 7px 6px; border-bottom: 1px solid var(--stripe); }
  tbody tr:nth-child(odd) { background: var(--stripe); }
  tbody tr:hover { background: #eef5fe; }
  .mono { font-family: ui-monospace, Consolas, monospace; }
  .muted { color: var(--muted); }
  .photo { max-width: 100%; border-radius: 8px; margin-top: 10px; }
  .empty { text-align: center; color: var(--muted); padding: 22px 0; }
  details summary { cursor: pointer; color: var(--primary); }
  footer { text-align: center; color: var(--muted); font-size: 12px; padding: 8px 0 24px; }
  @media (max-width: 560px) {
    body { padding: 8px; } .card { padding: 12px; } .result .model { font-size: 22px; }
    th, td { padding: 6px 4px; font-size: 13px; }
  }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>__TITLE__</h1>
    <p>__SUBTITLE__</p>
  </header>
  <div class="accent"></div>

  <div class="card">
    <h2>五位自编号查询</h2>
    <div class="row">
      <input id="number" class="grow" inputmode="numeric" maxlength="10" placeholder="例如 20000">
      <button class="primary" id="go">查询</button>
    </div>
    <div class="result" id="result"><div class="hint">输入五位自编号，边输入边出结果</div></div>
  </div>

  <div class="card">
    <h2>车型目录</h2>
    <div class="row" style="margin-bottom:10px">
      <input id="search" class="grow" placeholder="车型代号 / 厂商 / 俗称">
      <select id="fuel"></select>
      <select id="length"></select>
      <select id="body"></select>
    </div>
    <div id="catalog"></div>
  </div>

  <div class="card">
    <h2>地铁列车编号</h2>
    <div class="row" style="margin-bottom:10px">
      <input id="metro" class="grow" placeholder="例如 01001">
    </div>
    <div id="metro-result"><div class="empty">输入列车编号查询所属线路与车型</div></div>
  </div>

  <div class="card" id="spots-card">
    <h2>我的拍车记录 <span class="muted" id="spots-summary" style="font-weight:400"></span></h2>
    <div class="row" style="margin-bottom:10px">
      <input id="spot-search" class="grow" placeholder="按自编号 / 车型 / 备注筛选">
    </div>
    <div id="spots"></div>
  </div>

  <div class="card">
    <h2>数据概况</h2>
    <div id="stats" class="muted"></div>
    __NOTICES__
  </div>

  <footer>导出时间 __GENERATED__ · 数据来自北京公交车型查询器 · 本文件离线可用</footer>
</div>

<script>
const DATA = __DATA__;

function fmt(n, width) { return String(n).padStart(width || 5, '0'); }

function findRange(number) {
  for (const item of DATA.ranges) {
    if (number >= item[0] && number <= item[1]) return item;
  }
  return null;
}

function tagsOf(model) {
  const tag = DATA.tags[model] || ['未标注', '未标注', '未标注', '未标注'];
  return { manufacturer: tag[0], fuel: tag[1], length: tag[2], body: tag[3] || '未标注' };
}

function rangeSummary(model) {
  let segments = 0, covered = 0;
  for (const item of DATA.ranges) {
    if (item[2] === model) { segments += 1; covered += item[1] - item[0] + 1; }
  }
  return { segments, covered };
}

function showNumber(raw) {
  const box = document.getElementById('result');
  const value = raw.trim();
  if (value === '') {
    box.innerHTML = '<div class="hint">输入五位自编号，边输入边出结果</div>';
    return;
  }
  if (!/^[0-9０-９]{5}$/.test(value)) {
    box.innerHTML = '<div class="bad">格式不符</div><div class="hint">请输入五位数字自编号</div>';
    return;
  }
  const digits = value.replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  const number = parseInt(digits, 10);
  const hit = findRange(number);
  if (!hit) {
    box.innerHTML = '<div class="missing">未收录</div><div class="hint">暂无对应车型，请核对自编号</div>';
    return;
  }
  const model = hit[2];
  const tag = tagsOf(model);
  const summary = rangeSummary(model);
  const profile = DATA.profiles[model] || {};
  let html = '<div class="model">' + model + '</div>';
  html += '<div class="chips"><span class="chip">' + tag.manufacturer + '</span><span class="chip">' + tag.fuel +
          '</span><span class="chip">' + tag.length + '</span><span class="chip">' + tag.body + '</span></div>';
  html += '<div class="hint">属于 ' + fmt(hit[0]) + '-' + fmt(hit[1]) + ' · 该车型 ' + summary.segments + ' 个号段 / 覆盖 ' +
          summary.covered + ' 个编号</div>';
  if (profile.intro) html += '<div class="hint">' + profile.intro + '</div>';
  if (profile.image) html += '<img class="photo" src="' + profile.image + '" alt="' + model + '">';
  box.innerHTML = html;
}

let sortKey = 'model', sortAsc = true;
function renderCatalog() {
  const query = document.getElementById('search').value.trim().toLowerCase();
  const fuel = document.getElementById('fuel').value;
  const length = document.getElementById('length').value;
  const body = document.getElementById('body').value;
  const models = Object.keys(DATA.tags);
  const rows = [];
  for (const model of models) {
    const tag = tagsOf(model);
    const alias = (DATA.aliases[model] || []).join(' ');
    const haystack = (model + ' ' + tag.manufacturer + ' ' + alias).toLowerCase();
    if (query && haystack.indexOf(query) === -1) continue;
    if (fuel && tag.fuel !== fuel) continue;
    if (length && tag.length !== length) continue;
    if (body && tag.body !== body) continue;
    const summary = rangeSummary(model);
    rows.push({ model: model, manufacturer: tag.manufacturer, fuel: tag.fuel, length: tag.length,
                body: tag.body, segments: summary.segments, covered: summary.covered });
  }
  const numeric = sortKey === 'segments' || sortKey === 'covered';
  rows.sort((a, b) => {
    const left = a[sortKey], right = b[sortKey];
    const result = numeric ? left - right : String(left).localeCompare(String(right), 'zh');
    return sortAsc ? result : -result;
  });
  const box = document.getElementById('catalog');
  if (!rows.length) { box.innerHTML = '<div class="empty">没有匹配的车型</div>'; return; }
  let html = '<table><thead><tr>' +
    ['model:车型', 'manufacturer:厂商', 'fuel:燃料', 'length:车长', 'body:车体', 'segments:号段', 'covered:编号数']
      .map(pair => {
        const parts = pair.split(':');
        return '<th data-key="' + parts[0] + '">' + parts[1] + (sortKey === parts[0] ? (sortAsc ? ' ▲' : ' ▼') : '') + '</th>';
      }).join('') + '</tr></thead><tbody>';
  for (const row of rows) {
    html += '<tr><td class="mono">' + row.model + '</td><td>' + row.manufacturer + '</td><td>' + row.fuel +
            '</td><td>' + row.length + '</td><td>' + row.body + '</td><td>' + row.segments + '</td><td>' + row.covered + '</td></tr>';
  }
  box.innerHTML = html + '</tbody></table><div class="muted" style="margin-top:8px">共 ' + rows.length + ' 种车型</div>';
  box.querySelectorAll('th').forEach(th => th.addEventListener('click', () => {
    const key = th.getAttribute('data-key');
    if (key === sortKey) sortAsc = !sortAsc; else { sortKey = key; sortAsc = true; }
    renderCatalog();
  }));
}

function pad(value) { return String(value).padStart(5, '0'); }

function renderMetro(raw) {
  const box = document.getElementById('metro-result');
  const value = raw.trim().toUpperCase();
  if (!value) { box.innerHTML = '<div class="empty">输入列车编号查询所属线路与车型</div>'; return; }
  const key = pad(value);
  const hits = [];
  for (const line of Object.keys(DATA.metro)) {
    for (const item of DATA.metro[line]) {
      // 号码是定宽字符串，等长比较等价于数值比较，也能兼容带字母前缀的线路
      if (key >= pad(item[0]) && key <= pad(item[1])) hits.push([line, item]);
    }
  }
  if (!hits.length) { box.innerHTML = '<div class="empty">没有收录这个编号</div>'; return; }
  box.innerHTML = '<table><thead><tr><th>线路</th><th>号段</th><th>车型</th></tr></thead><tbody>' +
    hits.map(hit => '<tr><td>' + hit[0] + '</td><td class="mono">' + pad(hit[1][0]) + '-' + pad(hit[1][1]) +
      '</td><td class="mono">' + hit[1][2] + '</td></tr>').join('') + '</tbody></table>';
}

function renderSpots() {
  const box = document.getElementById('spots');
  const card = document.getElementById('spots-card');
  const rows = DATA.spots || [];
  if (!rows.length) { card.style.display = 'none'; return; }
  document.getElementById('spots-summary').textContent = DATA.spot_summary || '';
  const query = document.getElementById('spot-search').value.trim().toLowerCase();
  const shown = rows.filter(row =>
    !query || (row.number + ' ' + row.model + ' ' + row.note + ' ' + row.time).toLowerCase().indexOf(query) !== -1);
  if (!shown.length) { box.innerHTML = '<div class="empty">没有匹配的记录</div>'; return; }
  let html = '<table><thead><tr><th>时间</th><th>自编号</th><th>车型</th><th>备注</th></tr></thead><tbody>';
  for (const row of shown) {
    html += '<tr><td class="mono">' + row.time + '</td><td class="mono">' + row.number + '</td><td class="mono">' +
      row.model + '</td><td>' + (row.note || '') +
      (row.photo ? '<img class="photo" style="max-width:220px;display:block;margin-top:6px" src="' + row.photo + '" alt="">' : '') +
      '</td></tr>';
  }
  box.innerHTML = html + '</tbody></table><div class="muted" style="margin-top:8px">显示 ' + shown.length + ' / ' + rows.length + ' 条</div>';
}

function fillSelect(id, values) {
  const select = document.getElementById(id);
  select.innerHTML = '<option value="">全部</option>' +
    values.map(value => '<option value="' + value + '">' + value + '</option>').join('');
}

document.addEventListener('DOMContentLoaded', () => {
  const fuels = [...new Set(Object.values(DATA.tags).map(tag => tag[1]))].sort();
  const lengths = [...new Set(Object.values(DATA.tags).map(tag => tag[2]))].sort((a, b) => {
    const da = parseInt(a, 10), db = parseInt(b, 10);
    if (isNaN(da) || isNaN(db)) return String(a).localeCompare(String(b), 'zh');
    return da - db;
  });
  const bodies = [...new Set(Object.values(DATA.tags).map(tag => tag[3] || '未标注'))].sort((a, b) => a.localeCompare(b, 'zh'));
  fillSelect('fuel', fuels);
  fillSelect('length', lengths);
  fillSelect('body', bodies);
  const numberBox = document.getElementById('number');
  numberBox.addEventListener('input', () => showNumber(numberBox.value));
  document.getElementById('go').addEventListener('click', () => showNumber(numberBox.value));
  numberBox.addEventListener('keydown', event => { if (event.key === 'Enter') showNumber(numberBox.value); });
  document.getElementById('search').addEventListener('input', renderCatalog);
  document.getElementById('fuel').addEventListener('change', renderCatalog);
  document.getElementById('length').addEventListener('change', renderCatalog);
  document.getElementById('body').addEventListener('change', renderCatalog);
  const metroBox = document.getElementById('metro');
  metroBox.addEventListener('input', () => renderMetro(metroBox.value));
  document.getElementById('spot-search').addEventListener('input', renderSpots);
  renderCatalog();
  renderMetro('');
  renderSpots();
  numberBox.focus();
});
</script>
</body>
</html>
"""


MODEL_CARD_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__MODEL__ · 车型卡片</title>
<style>
  :root { --bg:#f0f3f8; --surface:#fff; --border:#dce3ec; --text:#16212e; --muted:#6b7b8f; --primary:#1668c4; --soft:#e8f1fc; }
  * { box-sizing: border-box; }
  body { margin:0; padding:18px; background:var(--bg); color:var(--text);
         font:15px/1.7 "Microsoft YaHei UI","Microsoft YaHei","PingFang SC",system-ui,sans-serif; }
  .card { max-width: 820px; margin:0 auto; background:var(--surface); border:1px solid var(--border); border-radius:12px; overflow:hidden; }
  .head { background:#123a66; color:#fff; padding:16px 20px; }
  .head h1 { margin:0; font-size:24px; word-break:break-all; }
  .head p { margin:6px 0 0; font-size:12px; color:#a9c6e4; }
  .accent { height:3px; background:#2e7bd6; }
  .body { padding:18px 20px 22px; }
  .chips { display:flex; gap:6px; flex-wrap:wrap; margin-bottom:14px; }
  .chip { background:var(--soft); color:var(--primary); border-radius:20px; padding:3px 12px; font-size:13px; font-weight:600; }
  h2 { font-size:13px; color:var(--muted); margin:18px 0 8px; letter-spacing:.5px; }
  .photos { display:flex; gap:10px; flex-wrap:wrap; }
  .photos figure { margin:0; }
  .photos img { max-width:100%; border-radius:8px; border:1px solid var(--border); }
  .photos figcaption { font-size:12px; color:var(--muted); margin-top:4px; }
  .text { white-space:pre-wrap; background:#f7fafd; border:1px solid var(--border); border-radius:8px; padding:10px 12px; }
  table { width:100%; border-collapse:collapse; font-size:14px; }
  th { text-align:left; color:var(--muted); font-size:12px; padding:8px 6px; border-bottom:1px solid var(--border); }
  td { padding:6px; border-bottom:1px solid #f2f5f9; font-family:ui-monospace,Consolas,monospace; }
  tbody tr:nth-child(odd) { background:#f7fafd; }
  .muted { color:var(--muted); font-size:12px; }
  footer { margin-top:20px; text-align:center; color:var(--muted); font-size:12px; }
  @media print { body { background:#fff; padding:0; } .card { border:0; } }
</style>
</head>
<body>
<div class="card">
  <div class="head">
    <h1>__MODEL__</h1>
    <p>__SUBTITLE__</p>
  </div>
  <div class="accent"></div>
  <div class="body">
    <div class="chips">__CHIPS__</div>
    __PHOTOS__
    __INTRODUCTION__
    __SPECIFICATIONS__
    <h2>对应号段</h2>
    __RANGES__
  </div>
</div>
<footer>导出时间 __GENERATED__ · 来自北京公交车型查询器</footer>
</body>
</html>
"""


def _chips(values: list[str]) -> str:
    return "".join(f'<span class="chip">{value}</span>' for value in values if value)


def render_model_card(payload: dict) -> str:
    """渲染单车型卡片（无脚本，可直接打印或发出去）。"""
    photos = payload.get("photos") or []
    if photos:
        figures = "".join(
            f'<figure><img src="{item["src"]}" alt="{item.get("caption", "")}">'
            + (f'<figcaption>{item["caption"]}</figcaption>' if item.get("caption") else "")
            + "</figure>"
            for item in photos
        )
        photos_html = f'<h2>图片</h2><div class="photos">{figures}</div>'
    else:
        photos_html = ""
    introduction = payload.get("introduction") or ""
    specifications = payload.get("specifications") or ""
    intro_html = f'<h2>简介</h2><div class="text">{introduction}</div>' if introduction else ""
    spec_html = f'<h2>技术参数</h2><div class="text">{specifications}</div>' if specifications else ""
    ranges = payload.get("ranges") or []
    if ranges:
        rows = "".join(
            f'<tr><td>{item["start"]}-{item["end"]}</td><td>{item["count"]}</td><td>{item.get("model", "")}</td></tr>'
            for item in ranges
        )
        ranges_html = (
            "<table><thead><tr><th>号段</th><th>编号数</th><th>车型</th></tr></thead><tbody>"
            + rows
            + "</tbody></table>"
            + f'<p class="muted">共 {len(ranges)} 个号段 · {payload.get("covered", 0)} 个自编号</p>'
        )
    else:
        ranges_html = '<p class="muted">当前 Excel 中没有该车型的号段。</p>'
    html = MODEL_CARD_TEMPLATE.replace("__MODEL__", str(payload.get("model", "")))
    html = html.replace("__SUBTITLE__", str(payload.get("subtitle", "")))
    html = html.replace("__CHIPS__", _chips([str(value) for value in payload.get("tags", [])]))
    html = html.replace("__PHOTOS__", photos_html)
    html = html.replace("__INTRODUCTION__", intro_html)
    html = html.replace("__SPECIFICATIONS__", spec_html)
    html = html.replace("__RANGES__", ranges_html)
    return html.replace("__GENERATED__", str(payload.get("generated", "")))


def write_model_card(path: Path, payload: dict) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(render_model_card(payload), encoding="utf-8")
    return path


def thumbnail_data_uri(path: Path, max_size: tuple[int, int] = MAX_THUMBNAIL, quality: int = THUMBNAIL_QUALITY) -> str:
    """把一张图片压成可内联的 data URI（读不到就返回空串）。"""
    try:
        from PIL import Image
    except ImportError:
        return ""
    try:
        with Image.open(path) as image:
            picture = image.convert("RGB")
            picture.thumbnail(max_size, Image.Resampling.LANCZOS)
            buffer = io.BytesIO()
            picture.save(buffer, format="JPEG", quality=quality, optimize=True)
    except Exception:
        return ""
    return "data:image/jpeg;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")


def render_html(payload: dict) -> str:
    """把 payload 渲染成完整的单文件 HTML。"""
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    # </script> 出现在字符串里会提前结束脚本块，转义掉更安全
    data = data.replace("</", "<\\/")
    notices = payload.get("notices") or []
    notice_html = ""
    if notices:
        items = "".join(f"<li>{text}</li>" for text in notices)
        notice_html = f'<details style="margin-top:8px"><summary>数据提醒（{len(notices)} 条）</summary><ul>{items}</ul></details>'
    html = HTML_TEMPLATE.replace("__DATA__", data)
    html = html.replace("__TITLE__", str(payload.get("title", "北京公交车型查询器")))
    html = html.replace("__SUBTITLE__", str(payload.get("subtitle", "")))
    html = html.replace("__GENERATED__", str(payload.get("generated", "")))
    return html.replace("__NOTICES__", notice_html)


def write_html(path: Path, payload: dict) -> Path:
    """写出网页版；返回实际写入的路径。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(render_html(payload), encoding="utf-8")
    return path
