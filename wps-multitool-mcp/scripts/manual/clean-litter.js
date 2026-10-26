/**
 * 一次性清理：关掉历次探测留下的"空白未保存"文档。
 * 只关这种：名字是 WPS 默认名（文字文稿N / 工作簿N / 演示文稿N）且从未存过盘（Path 为空）。
 * 用户真实文档（有路径、或自定义名）一个都不碰。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-clean-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '60000';
const path = require('path');
const { wpsClient } = require(path.resolve(__dirname, '..', '..', 'dist', 'client', 'wps-client.js'));

(async () => {
  const r = await wpsClient.executeMethod('getOpenDocuments', {}, 'word');
  const docs = (r.data && r.data.documents) || [];
  let closed = 0;
  for (const d of docs) {
    const nm = String(d.name || '');
    if (!/^文字文稿\d+$/.test(nm.replace(/\.docx?$/i, ''))) continue;
    const c = await wpsClient.executeMethod('closeDocument', { name: nm, saveChanges: false }, 'word')
      .catch((e) => ({ success: false, error: String(e) }));
    if (c.success) closed++;
    console.log(`word ${nm}: ${c.success ? '已关' : '失败 ' + (c.error || '')}`);
  }

  const eb = await wpsClient.executeMethod('getOpenWorkbooks', {}, 'excel');
  for (const b of (eb.data && eb.data.workbooks) || []) {
    const nm = String(b.name || '');
    if (!/^工作簿\d+$/.test(nm.replace(/\.xlsx?$/i, ''))) continue;
    const c = await wpsClient.executeMethod('closeWorkbook', { name: nm, saveChanges: false }, 'excel')
      .catch((e) => ({ success: false, error: String(e) }));
    if (c.success) closed++;
    console.log(`excel ${nm}: ${c.success ? '已关' : '失败 ' + (c.error || '')}`);
  }

  const pp = await wpsClient.executeMethod('getOpenPresentations', {}, 'ppt');
  for (const p of (pp.data && pp.data.presentations) || []) {
    const nm = String(p.name || '');
    if (!/^演示文稿\d+$/.test(nm)) continue;
    const c = await wpsClient.executeMethod('closePresentation', { name: nm, saveChanges: false }, 'ppt')
      .catch((e) => ({ success: false, error: String(e) }));
    if (c.success) closed++;
    console.log(`ppt ${nm}: ${c.success ? '已关' : '失败 ' + (c.error || '')}`);
  }
  console.log(`共关掉 ${closed} 个空白文档`);
  process.exit(0);
})();
