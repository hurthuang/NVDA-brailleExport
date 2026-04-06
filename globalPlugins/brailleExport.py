# -*- coding: utf-8 -*-
# brailleExport — NVDA Global Plugin
# 攔截 NVDA 點字輸出並匯出為檔案

import globalPluginHandler
import braille
import ui
import wx
import os
import datetime
import scriptHandler
from scriptHandler import script
import addonHandler
import gui
import config
import logHandler
import api  # 引入 NVDA 的 api 模組以使用剪貼簿功能

addonHandler.initTranslation()
log = logHandler.log

# ──────────────────────────────────────────────
# Cell 轉換
# ──────────────────────────────────────────────
BRAILLE_UNICODE_BASE = 0x2800

def cells_to_unicode(cells):
    """NVDA cells → Unicode 點字字串（保留完整的 8 點資訊）"""
    return "".join(chr(BRAILLE_UNICODE_BASE | (c & 0xFF)) for c in cells)

# NABCC (North American Braille Computer Code) lookup table
_NABCC = {
    0x01: 'a', 0x02: '1', 0x03: 'b', 0x04: "'", 0x05: 'k',
    0x06: '2', 0x07: 'l', 0x08: '`', 0x09: 'c', 0x0A: 'i', 0x0B: 'f',
    0x0C: '/', 0x0D: 'm', 0x0E: 's', 0x0F: 'p', 0x10: '"', 0x11: 'e',
    0x12: '3', 0x13: 'h', 0x14: '9', 0x15: 'o', 0x16: '6', 0x17: 'r',
    0x18: '^', 0x19: 'd', 0x1A: 'j', 0x1B: 'g', 0x1C: '>', 0x1D: 'n',
    0x1E: 't', 0x1F: 'q', 0x20: ',', 0x21: '*', 0x22: '5', 0x23: '<',
    0x24: '-', 0x25: 'u', 0x26: '8', 0x27: 'v', 0x28: '.', 0x29: '%',
    0x2A: '[', 0x2B: '$', 0x2C: '+', 0x2D: 'x', 0x2E: '!', 0x2F: '&',
    0x30: ';', 0x31: ':', 0x32: '4', 0x33: '\\', 0x34: '0', 0x35: 'z',
    0x36: '7', 0x37: '(', 0x38: '_', 0x39: '?', 0x3A: 'w', 0x3B: ']',
    0x3C: '#', 0x3D: 'y', 0x3E: ')', 0x3F: '=',
}

def cells_to_brf(cells):
    """NVDA cells → BRF ASCII (BRF 本身為 6 點標準)"""
    parts = []
    for c in cells:
        bits = c & 0x3F
        if bits == 0:
            parts.append(' ')   
        else:
            parts.append(_NABCC.get(bits, '?'))
    return ''.join(parts).rstrip(' ')

def _default_export_dir():
    return os.path.join(os.path.expanduser("~"), "Desktop")

# ──────────────────────────────────────────────
# 設定存取
# ──────────────────────────────────────────────
CONFIG_SECTION = "brailleExport"
CONFIG_DEFAULTS = {
    "exportDir": _default_export_dir(),
    "exportFormat": "unicode",
    "exportDestination": "clipboard",  # 已將預設修改為剪貼簿 (clipboard)
    "addTimestamp": "True",
    "cellsPerLine": "40",
}

def _cfg():
    if CONFIG_SECTION not in config.conf:
        config.conf[CONFIG_SECTION] = {}
    sec = config.conf[CONFIG_SECTION]
    for k, v in CONFIG_DEFAULTS.items():
        if k not in sec:
            sec[k] = v
    return sec

def _cfg_bool(key):
    return str(_cfg().get(key, CONFIG_DEFAULTS[key])).lower() in ("true", "1", "yes")

def _cfg_int(key):
    try:
        return int(_cfg().get(key, CONFIG_DEFAULTS[key]))
    except (ValueError, TypeError):
        return int(CONFIG_DEFAULTS[key])

def _cfg_str(key):
    return str(_cfg().get(key, CONFIG_DEFAULTS[key]))


# ══════════════════════════════════════════════
# 主外掛類別
# ══════════════════════════════════════════════
class GlobalPlugin(globalPluginHandler.GlobalPlugin):

    scriptCategory = "點字輸出匯出"

    def __init__(self):
        super().__init__()
        self._recording = False
        self._buffer = []        
        self._origUpdate = None
        self._hooked = False
        self._hookBraille()
        self._buildMenu()
        log.info("[brailleExport] loaded OK (Default: Clipboard)")

    def _hookBraille(self):
        try:
            handler = braille.handler
            if handler is None:
                return
            
            if hasattr(handler, "_writeCells") and not getattr(handler, "_beExportHooked", False):
                self._origWriteCells = handler._writeCells
                handler._writeCells = self._hookedWriteCells
                handler._beExportHooked = True
                self._hookType = "writeCells"
                self._hooked = True
            elif hasattr(handler, "update") and not getattr(handler, "_beExportHooked", False):
                self._origUpdate = handler.update
                handler.update = self._hookedUpdate
                handler._beExportHooked = True
                self._hookType = "update"
                self._hooked = True
        except Exception as e:
            log.error(f"[brailleExport] hookBraille error: {e}")

    def _ensureHook(self):
        if not self._hooked:
            self._hookBraille()

    def _process_new_frame(self, cl):
        if not self._buffer:
            self._buffer.append(cl)
            return

        prev = self._buffer[-1]
        if cl == prev:
            return

        if len(cl) == len(prev):
            diffs = [(i, cl[i], prev[i]) for i in range(len(cl)) if cl[i] != prev[i]]
            if len(diffs) == 1:
                i, c_new, c_old = diffs[0]
                if (c_new & 0x3F) == (c_old & 0x3F):
                    if bin(c_new).count('1') < bin(c_old).count('1'):
                        self._buffer[-1] = cl  
                    return 
        
        self._buffer.append(cl)

    def _hookedWriteCells(self, cells):
        if self._recording and cells:
            self._process_new_frame(list(cells))
        self._origWriteCells(cells)

    def _hookedUpdate(self):
        self._origUpdate()
        if self._recording:
            try:
                cl = list(braille.handler.buffer.windowBrailleCells)
                if cl:
                    self._process_new_frame(cl)
            except Exception:
                pass

    def _unhookBraille(self):
        try:
            handler = braille.handler
            if handler is None:
                return
            if getattr(handler, "_beExportHooked", False):
                if hasattr(self, "_origWriteCells") and self._origWriteCells:
                    handler._writeCells = self._origWriteCells
                    self._origWriteCells = None
                elif self._origUpdate:
                    handler.update = self._origUpdate
                    self._origUpdate = None
                handler._beExportHooked = False
            self._hooked = False
        except Exception as e:
            pass

    def _buildMenu(self):
        try:
            self._menu = wx.Menu()
            item = self._menu.Append(wx.ID_ANY, "立即匯出目前畫面(&E)")
            gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self._onExportCurrent, item)
            item = self._menu.Append(wx.ID_ANY, "開始錄製(&R)")
            gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self._onStartRecord, item)
            item = self._menu.Append(wx.ID_ANY, "停止錄製並匯出(&S)")
            gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self._onStopExport, item)
            self._menu.AppendSeparator()
            item = self._menu.Append(wx.ID_ANY, "設定(&O)…")
            gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self._onSettings, item)
            self._mainMenuItem = gui.mainFrame.sysTrayIcon.menu.Insert(
                2, wx.ID_ANY, "點字匯出(&X)", self._menu
            )
        except Exception as e:
            log.error(f"[brailleExport] buildMenu error: {e}")

    @script(description="立即匯出目前點字顯示畫面快照", gesture="kb:NVDA+shift+c")
    def script_exportCurrent(self, gesture):
        self._exportCurrentCells()

    @script(description="開始錄製點字輸出（再按一次停止並匯出）", gesture="kb:NVDA+shift+r")
    def script_toggleRecord(self, gesture):
        if self._recording:
            self._stopRecordAndExport()
        else:
            self._startRecord()

    @script(description="開啟點字匯出設定對話框")
    def script_openSettings(self, gesture):
        wx.CallAfter(self._openSettingsDlg)

    def _onExportCurrent(self, evt):
        wx.CallAfter(self._exportCurrentCells)

    def _onStartRecord(self, evt):
        wx.CallAfter(self._startRecord)

    def _onStopExport(self, evt):
        wx.CallAfter(self._stopRecordAndExport)

    def _onSettings(self, evt):
        wx.CallAfter(self._openSettingsDlg)

    def _openSettingsDlg(self):
        gui.mainFrame.prePopup()
        dlg = BrailleExportSettingsDialog(gui.mainFrame)
        dlg.ShowModal()
        dlg.Destroy()
        gui.mainFrame.postPopup()

    def _exportCurrentCells(self):
        self._ensureHook()
        try:
            handler = braille.handler
            if handler is None:
                ui.message("目前沒有點字處理器。")
                return
            cells = list(handler.buffer.windowBrailleCells)
            
            try:
                buf = handler.buffer
                if hasattr(buf, 'cursorPos') and hasattr(buf, 'windowPos'):
                    if buf.cursorPos is not None and buf.windowPos is not None:
                        cursor_idx = buf.cursorPos - buf.windowPos
                        if 0 <= cursor_idx < len(cells):
                            cells[cursor_idx] &= 0x3F
            except Exception:
                pass
                
            if not any(c != 0 for c in cells):
                ui.message("目前點字顯示器為空白。")
                return
                
            text = self._generateText([cells])
            dest = _cfg_str("exportDestination")
            
            if dest == "clipboard":
                if api.copyToClip(text.strip()):
                    ui.message("快照已複製到剪貼簿。")
                else:
                    ui.message("複製到剪貼簿失敗。")
            else:
                path = self._buildPath("snapshot")
                self._writeFile(path, text)
                ui.message(f"已匯出：{path}")
                
        except Exception as e:
            ui.message(f"處理失敗：{e}")

    def _startRecord(self):
        self._ensureHook()
        if self._recording:
            ui.message("錄製已在進行中。")
            return
        self._buffer.clear()
        ui.message("開始錄製點字輸出。再次執行指令停止並匯出。")
        wx.CallLater(400, self._beginRecordingNow)

    def _beginRecordingNow(self):
        self._buffer.clear()
        self._recording = True
        log.info("[brailleExport] recording started")

    def _stopRecordAndExport(self):
        if not self._recording:
            ui.message("目前沒有進行錄製。")
            return
        self._recording = False
        if not self._buffer:
            ui.message("緩衝區為空，未產生內容。")
            return
            
        n = len(self._buffer)
        text = self._generateText(self._buffer)
        self._buffer.clear()
        
        dest = _cfg_str("exportDestination")
        if dest == "clipboard":
            if api.copyToClip(text.strip()):
                ui.message(f"錄製完成，共 {n} 幀，已複製到剪貼簿。")
            else:
                ui.message("複製到剪貼簿失敗。")
        else:
            path = self._buildPath("record")
            self._writeFile(path, text)
            ui.message(f"錄製完成，共 {n} 幀，已匯出：{path}")

    # ──────────────────────────────────────────
    # 產生文字與檔案
    # ──────────────────────────────────────────
    def _generateText(self, frames):
        """將紀錄的 frames 轉換為完整的格式化文字"""
        fmt = _cfg_str("exportFormat")
        cells_per_line = _cfg_int("cellsPerLine")
        lines = []
        for frame in frames:
            raw = cells_to_unicode(frame) if fmt == "unicode" else cells_to_brf(frame)
            for i in range(0, max(len(raw), 1), cells_per_line):
                chunk = raw[i:i + cells_per_line]
                if chunk.strip("\u2800 "):
                    lines.append(chunk)
        return "\n".join(lines) + "\n"

    def _buildPath(self, prefix):
        export_dir = _cfg_str("exportDir")
        fmt = _cfg_str("exportFormat")
        add_ts = _cfg_bool("addTimestamp")
        ts = f"_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}" if add_ts else ""
        ext = ".txt" if fmt == "unicode" else ".brf"
        os.makedirs(export_dir, exist_ok=True)
        return os.path.join(export_dir, f"braille_{prefix}{ts}{ext}")

    def _writeFile(self, path, content):
        """將文字寫入實體檔案"""
        fmt = _cfg_str("exportFormat")
        enc = "utf-8" if fmt == "unicode" else "ascii"
        with open(path, "w", encoding=enc, newline="\r\n") as f:
            f.write(content)
        log.info(f"[brailleExport] written to {path}")

    def terminate(self):
        self._recording = False
        self._unhookBraille()
        try:
            gui.mainFrame.sysTrayIcon.menu.Remove(self._mainMenuItem)
        except Exception:
            pass
        super().terminate()


# ══════════════════════════════════════════════
# 設定對話框
# ══════════════════════════════════════════════
class BrailleExportSettingsDialog(wx.Dialog):

    def __init__(self, parent):
        super().__init__(parent, title="點字輸出匯出 — 設定")
        sizer = wx.BoxSizer(wx.VERTICAL)

        # ── 輸出目標 ──
        destBox = wx.StaticBox(self, label="輸出目標")
        destSizer = wx.StaticBoxSizer(destBox, wx.HORIZONTAL)
        dest = _cfg_str("exportDestination")
        self._radioFile = wx.RadioButton(self, label="儲存成檔案", style=wx.RB_GROUP)
        self._radioClipboard = wx.RadioButton(self, label="複製到剪貼簿")
        self._radioFile.SetValue(dest != "clipboard")
        self._radioClipboard.SetValue(dest == "clipboard")
        destSizer.Add(self._radioFile, 0, wx.ALL, 4)
        destSizer.Add(self._radioClipboard, 0, wx.ALL, 4)
        sizer.Add(destSizer, 0, wx.EXPAND | wx.ALL, 8)

        # ── 匯出目錄 ──
        dirBox = wx.StaticBox(self, label="匯出目錄 (僅存成檔案時有效)")
        dirSizer = wx.StaticBoxSizer(dirBox, wx.HORIZONTAL)
        self._dirCtrl = wx.TextCtrl(self, value=_cfg_str("exportDir"), size=(320, -1))
        browseBtn = wx.Button(self, label="瀏覽(&B)…")
        browseBtn.Bind(wx.EVT_BUTTON, self._onBrowse)
        dirSizer.Add(self._dirCtrl, 1, wx.EXPAND | wx.ALL, 4)
        dirSizer.Add(browseBtn, 0, wx.ALL, 4)
        sizer.Add(dirSizer, 0, wx.EXPAND | wx.ALL, 8)

        # ── 格式 ──
        fmtBox = wx.StaticBox(self, label="文字格式")
        fmtSizer = wx.StaticBoxSizer(fmtBox, wx.VERTICAL)
        fmt = _cfg_str("exportFormat")
        self._radioUnicode = wx.RadioButton(self, label="Unicode 點字文字  ⠃⠗⠁⠊⠇⠇⠑", style=wx.RB_GROUP)
        self._radioBrf = wx.RadioButton(self, label="BRF ASCII 點字  可送點字印表")
        self._radioUnicode.SetValue(fmt == "unicode")
        self._radioBrf.SetValue(fmt == "brf")
        fmtSizer.Add(self._radioUnicode, 0, wx.ALL, 4)
        fmtSizer.Add(self._radioBrf, 0, wx.ALL, 4)
        sizer.Add(fmtSizer, 0, wx.EXPAND | wx.ALL, 8)

        # ── 時間戳記 ──
        self._tsCheck = wx.CheckBox(self, label="檔名加入時間戳記(&T)")
        self._tsCheck.SetValue(_cfg_bool("addTimestamp"))
        sizer.Add(self._tsCheck, 0, wx.ALL, 8)

        # ── 每行格數 ──
        rowSizer = wx.BoxSizer(wx.HORIZONTAL)
        rowSizer.Add(wx.StaticText(self, label="每行方數(&C)："), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)
        self._cellsSpin = wx.SpinCtrl(self, min=10, max=200, initial=_cfg_int("cellsPerLine"))
        rowSizer.Add(self._cellsSpin, 0)
        sizer.Add(rowSizer, 0, wx.ALL, 8)

        btnSizer = wx.StdDialogButtonSizer()
        okBtn = wx.Button(self, wx.ID_OK, "確定")
        okBtn.Bind(wx.EVT_BUTTON, self._onOk)
        cancelBtn = wx.Button(self, wx.ID_CANCEL, "取消")
        btnSizer.AddButton(okBtn)
        btnSizer.AddButton(cancelBtn)
        btnSizer.Realize()
        sizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 8)

        self.SetSizer(sizer)
        sizer.Fit(self)
        self.CenterOnParent()
        okBtn.SetDefault()

    def _onBrowse(self, evt):
        dlg = wx.DirDialog(self, "選擇匯出目錄", self._dirCtrl.GetValue(), style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            self._dirCtrl.SetValue(dlg.GetPath())
        dlg.Destroy()

    def _onOk(self, evt):
        sec = _cfg()
        sec["exportDestination"] = "clipboard" if self._radioClipboard.GetValue() else "file"
        sec["exportDir"] = self._dirCtrl.GetValue()
        sec["exportFormat"] = "unicode" if self._radioUnicode.GetValue() else "brf"
        sec["addTimestamp"] = str(self._tsCheck.GetValue())
        sec["cellsPerLine"] = str(self._cellsSpin.GetValue())
        self.EndModal(wx.ID_OK)