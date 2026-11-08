#!/usr/bin/env bash
# WPS Linux 加载项安装脚本（多面手 MCP · Linux 线）
# 依据：B5-T4/B5-T5 云端真机 playbook（linux/B5T4-REPORT.md §4.1、B5T5-REPORT.md §3.3/§八）
# 用法：sudo bash install.sh [--dry-run]
# 前提：已安装 WPS Office Linux 12.1.2.x；本脚本旁须有 addon-linux 树（本仓库）。
# 红线：所有改动前备份（<file>.bak_multitool）；幂等（重复跑安全）；首次启动的三类弹窗需手动点选（见末尾指引）。
set -u
DRY=0; [ "${1:-}" = "--dry-run" ] && DRY=1
act() { if [ "$DRY" = "1" ]; then echo "  [dry-run] $1"; else eval "$1"; fi; }
bak() { # bak <file>：首次改动前备份
  local f="$1"; [ -f "$f" ] && [ ! -f "$f.bak_multitool" ] && cp "$f" "$f.bak_multitool" && echo "  备份: $f.bak_multitool"; return 0
}
echo "== WPS Linux 加载项安装（多面手 MCP）$([ $DRY = 1 ] && echo '[dry-run 预演]') =="

# ── 0. 找 WPS 安装目录与配置目录 ──
WPSDIR=$(find /opt /usr -maxdepth 4 -type d -name office6 -path '*kingsoft*' 2>/dev/null | head -1)
[ -z "$WPSDIR" ] && { echo "✗ 未找到 WPS office6 目录（先装 WPS Office Linux 12.1.2.x）"; exit 1; }
echo "WPS 目录: $WPSDIR"
CFG=$(find "$HOME/.local/share" "$HOME/.config" -maxdepth 3 -type d -name kingsoft-wps 2>/dev/null | head -1)
[ -z "$CFG" ] && { echo "✗ 未找到 WPS 用户配置目录（先启动一次 WPS 让它自建，再跑本脚本）"; exit 1; }
echo "配置目录: $CFG"

# ── 1. 版本体检 ──
VER=$(dpkg -l wps-office 2>/dev/null | awk '/^ii/{print $3}')
case "$VER" in
  12.1.2*) echo "版本: $VER ✓（12.1.2.x 同步门面已验证）" ;;
  "")      echo "✗ 未检测到 wps-office 包"; exit 1 ;;
  *)       echo "⚠ 版本 $VER 非 12.1.2.x——11.x 走异步兜底也能用，但同步门面未在此版本验证"; ;;
esac

# ── 2. 清理强杀遗留的单实例锁/IPC 残留（B5-T5：301 个残留会让新 WPS 假起不来）──
STALE=$(ls /tmp/__\#_PROME_INSTANCE_EXIST_KEY_#__*.lock /tmp/qipc_* /tmp/JSAPI_IPCJSBIND* 2>/dev/null | wc -l)
if [ "$STALE" -gt 0 ] && [ "$DRY" = "0" ]; then rm -f /tmp/__\#_PROME_INSTANCE_EXIST_KEY_#__*.lock /tmp/qipc_* /tmp/JSAPI_IPCJSBIND* 2>/dev/null; fi
echo "IPC 残留清理: $STALE 项"

# ── 3. oem.ini 补 JsApiPlugin=true（缺开关时不向通道发任何请求）──
OEM="$WPSDIR/oem.ini"
[ -f "$OEM" ] || { echo "✗ 未找到 $OEM"; exit 1; }
bak "$OEM"
if grep -q 'JsApiPlugin=true' "$OEM"; then echo "oem.ini: JsApiPlugin 已在 ✓"
else act "sed -i 's/^\[Common_\|^\[Common\]/&\\nJsApiPlugin=true/' '$OEM' 2>/dev/null; grep -q 'JsApiPlugin=true' '$OEM' || echo 'JsApiPlugin=true' >> '$OEM'"; echo "oem.ini: 追加 JsApiPlugin=true"; fi

# ── 4. NSS 5 库软链系统版（12.x 自带缺失时 CEF 崩）──
SYSNSS=/usr/lib/x86_64-linux-gnu
for lib in libnss3.so libnssutil3.so libsmime3.so libssl3.so libplc4.so; do
  if [ ! -e "$WPSDIR/$lib" ] && [ -f "$SYSNSS/$lib" ]; then act "ln -s '$SYSNSS/$lib' '$WPSDIR/$lib'"; echo "  软链: $lib"; fi
done

# ── 5. Office.conf 切非融合模式（B5-T5：残缺配置致命——只在完整文件上改，缺文件让 WPS 自建）──
CONF=$(find "$CFG" -maxdepth 2 -name Office.conf 2>/dev/null | head -1)
if [ -z "$CONF" ]; then
  echo "⚠ Office.conf 不存在：先启动一次 WPS 让它自建完整配置，再重跑本脚本（B5-T5：残缺配置会让 WPS exit=255）"
elif grep -q 'prome_independ' "$CONF"; then echo "Office.conf: 已是非融合模式 ✓"
else bak "$CONF"; act "sed -i 's/prome_fushion/prome_independ/' '$CONF'"; echo "Office.conf: 切 prome_independ"; fi

# ── 6. jsplugins.xml 在线格式（指向本机 3889 静态服务）──
JSP=$(find "$CFG" -maxdepth 2 -name jsplugins.xml 2>/dev/null | head -1)
[ -z "$JSP" ] && JSP="$CFG/jsplugins.xml"
bak "$JSP"
JSCONT='<jsplugins>
  <jspluginonline name="wps-multitool-wps" type="wps" url="http://127.0.0.1:3889/wps/"/>
  <jspluginonline name="wps-multitool-et" type="et" url="http://127.0.0.1:3889/et/"/>
  <jspluginonline name="wps-multitool-wpp" type="wpp" url="http://127.0.0.1:3889/wpp/"/>
</jsplugins>'
if [ "$DRY" = "1" ]; then echo "  [dry-run] 写 $JSP（三条 jspluginonline）"
else echo "$JSCONT" > "$JSP"; echo "jsplugins.xml: 写入三型在线入口 → 127.0.0.1:3889"; fi

# ── 7. 移走 publish.xml（install="null" 会让 WPS 忽略在线 URL）──
PUB=$(find "$CFG" -maxdepth 2 -name publish.xml 2>/dev/null | head -1)
if [ -n "$PUB" ]; then act "mv '$PUB' '${PUB}.bak_multitool'"; echo "publish.xml: 已移走（${PUB}.bak_multitool）"; else echo "publish.xml: 不在（✓）"; fi

# ── 8. 静态服务说明 ──
HERE=$(cd "$(dirname "$0")" && pwd)
cat <<TIP

== 安装完成。后续两步 ==

[1] 起静态服务（加载项文件源，端口 3889）：
      node "$HERE/serve3889.js" "$HERE" &
    （MCP Server 侧的通道 27740 由 npx wps-multitool-mcp 启动时自动拉起）

[2] 启动 WPS，手动处理三类首装弹窗（各应用一次）：
    ① EULA：先勾选「Have read and agree」复选框，Confirm 才可用（仅回车无效）
    ② 默认软件提示：回车跳过
    ③ 首次信任框：「Addin ... trust or not?」必须点 OK——不点则加载项永不握手
    三型全部 connected 即成功（可用主仓 27741 管理端查看）。

提示：三型各有一个 EULA 窗口、可能几何重叠（B5-T5 实测），拖开逐个处理。
TIP
