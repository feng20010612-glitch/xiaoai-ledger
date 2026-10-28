#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""小爱记账 - 完全独立版"""

import sys, os, json, time, urllib.parse, webbrowser, socket, threading, tempfile, zipfile, shutil, base64
import re
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
try:
    import openpyxl
    HAS_XL = True
except:
    HAS_XL = False

# ── 路径 ──
# 数据目录相对于exe/脚本自身位置，拷到U盘也能用
if getattr(sys, 'frozen', False):
    EXE_DIR = os.path.dirname(sys.executable)
    BUNDLE_DIR = sys._MEIPASS
else:
    EXE_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = EXE_DIR

DATA_DIR = os.path.join(EXE_DIR, 'data')
RECORDS_FILE = os.path.join(DATA_DIR, 'records.json')
SETTINGS_FILE = os.path.join(DATA_DIR, 'settings.json')
TRASH_FILE = os.path.join(DATA_DIR, 'trash.json')
BUDGET_FILE = os.path.join(DATA_DIR, 'budget.json')
RECURRING_FILE = os.path.join(DATA_DIR, 'recurring.json')
BACKUP_DIR = os.path.join(EXE_DIR, 'data_backup')
VERSION_FILE = os.path.join(DATA_DIR, 'version.json')
APP_VERSION = '2.15.27'

# 升级时只替换「程序文件」，绝不碰 data\（账本、预算、设置全保留）
UPGRADE_PROGRAM_FILES = [
    '记账本-v2.html', 'chart.umd.min.js', 'icon.ico',
    '小爱记账.py', '小爱记账.exe',
    '使用手册.md', '更新日志.md',
    '使用说明.txt', '启动.bat',
]

DEFAULT_EXPENSE = ["吃喝饮食","居住水电","交通出行","生活日用","医疗健康","服饰美容","数码电子","娱乐休闲","人情社交","学习提升","通讯网络","投资理财","投资亏损","储蓄存款","其他杂项","烂账"]
DEFAULT_INCOME  = ["工资","兼职外快","投资赎回","投资收益","储蓄取出","红包礼金","退款","其他收入","烂账"]

# 退款分类名（唯一真源：改名只改这里）
REFUND_CAT = '退款'

MIME = {
    '.html':'text/html;charset=utf-8',
    '.css':'text/css;charset=utf-8',
    '.js':'application/javascript;charset=utf-8',
    '.json':'application/json;charset=utf-8',
    '.md':'text/plain;charset=utf-8',
    '.png':'image/png',
    '.jpg':'image/jpeg',
    '.ico':'image/x-icon',
}

def read_version_file():
    """读版本信息，没有就返回当前程序版本"""
    try:
        if os.path.isfile(VERSION_FILE):
            with open(VERSION_FILE, 'r', encoding='utf-8') as f:
                v = json.load(f)
            if isinstance(v, dict) and v.get('version'):
                return v
    except Exception:
        pass
    return {'version': APP_VERSION, 'updatedAt': '', 'note': ''}


def write_version_file(ver, note=''):
    """写入版本号（升级成功后调用）"""
    ensure_dir()
    data = {'version': ver or APP_VERSION,
            'updatedAt': time.strftime('%Y-%m-%d %H:%M:%S'),
            'note': note or ''}
    try:
        with open(VERSION_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return data


def _ver_tuple(v):
    """'2.11.0' -> (2, 11, 0)"""
    out = []
    for part in str(v or '').split('.'):
        digits = ''.join(ch for ch in part if ch.isdigit())
        try:
            out.append(int(digits) if digits else 0)
        except Exception:
            out.append(0)
    return tuple(out) if out else (0,)


def compare_version(a, b):
    """a > b -> 1；a == b -> 0；a < b -> -1"""
    ta, tb = _ver_tuple(a), _ver_tuple(b)
    n = max(len(ta), len(tb))
    ta = ta + (0,) * (n - len(ta))
    tb = tb + (0,) * (n - len(tb))
    if ta > tb:
        return 1
    if ta < tb:
        return -1
    return 0


def ensure_dir(): os.makedirs(DATA_DIR, exist_ok=True)

def load_records():
    ensure_dir()
    if not os.path.exists(RECORDS_FILE): return []
    with open(RECORDS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_records(r):
    ensure_dir()
    with open(RECORDS_FILE, 'w', encoding='utf-8') as f:
        json.dump(r, f, ensure_ascii=False, indent=2)

def load_settings():
    ensure_dir()
    if not os.path.exists(SETTINGS_FILE):
        return {'expenseCats': list(DEFAULT_EXPENSE), 'incomeCats': list(DEFAULT_INCOME), 'customIcons': {}, 'subCats': {}}
    with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
        return validate_settings(json.load(f))

def save_settings(s):
    ensure_dir()
    with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(s, f, ensure_ascii=False, indent=2)

def validate_settings(s):
    """保证 settings 包含所有必要字段"""
    if 'expenseCats' not in s:
        s['expenseCats'] = list(DEFAULT_EXPENSE)
    if 'incomeCats' not in s:
        s['incomeCats'] = list(DEFAULT_INCOME)
    if 'customIcons' not in s:
        s['customIcons'] = {}
    if 'subCats' not in s or not isinstance(s.get('subCats'), dict):
        s['subCats'] = {}
    # 回收站保留天数：缺省 30，0 = 永久保留
    if 'trashDays' not in s:
        s['trashDays'] = TRASH_DEFAULT_DAYS
    else:
        try:
            _n = int(s.get('trashDays'))
            s['trashDays'] = _n if _n >= 0 else TRASH_DEFAULT_DAYS
        except (TypeError, ValueError):
            s['trashDays'] = TRASH_DEFAULT_DAYS
    return s

def load_trash():
    if not os.path.exists(TRASH_FILE): return []
    with open(TRASH_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_trash(t):
    ensure_dir()
    with open(TRASH_FILE, 'w', encoding='utf-8') as f:
        json.dump(t, f, ensure_ascii=False, indent=2)

# ── 回收站自动清理 ──
TRASH_DEFAULT_DAYS = 30   # 默认保留 30 天；0 = 永久保留

def trash_keep_days():
    """读设置里的回收站保留天数；没设置/非法就用默认 30。0 表示永久保留。"""
    try:
        v = load_settings().get('trashDays', TRASH_DEFAULT_DAYS)
        n = int(v)
        return n if n >= 0 else TRASH_DEFAULT_DAYS
    except (TypeError, ValueError):
        return TRASH_DEFAULT_DAYS

def _rec_deleted_at(rec):
    """取记录的删除时间（毫秒）。老数据没 deletedAt 就退回 createdAt。"""
    for k in ('deletedAt', 'createdAt'):
        try:
            v = rec.get(k)
            if v:
                v = int(v)
                if v > 0:
                    return v
        except (TypeError, ValueError):
            pass
    return None

def purge_old_trash(items=None):
    """按保留天数清掉过期记录，返回 (保留列表, 清掉条数)。
    items 传 None 时直接读文件。days<=0 表示永久保留，原样返回。"""
    days = trash_keep_days()
    if days <= 0:
        lst = load_trash() if items is None else items
        return lst, 0
    lst = load_trash() if items is None else list(items)
    now = int(time.time() * 1000)
    cutoff = now - days * 86400 * 1000
    kept, dropped, needs_stamp = [], 0, []
    for x in lst:
        t = _rec_deleted_at(x)
        if t is None:
            # 完全没有时间信息的老记录：补上当前时间，从今天起算保留期
            needs_stamp.append(x)
            kept.append(x)
        elif t < cutoff:
            dropped += 1
        else:
            kept.append(x)
    if needs_stamp:
        for x in needs_stamp:
            x['deletedAt'] = now
    return kept, dropped

def auto_purge_trash():
    """惰性清理：在读取回收站时调用。有变化才写盘，避免无意义 IO。"""
    lst = load_trash()
    if not lst:
        return 0
    kept, dropped = purge_old_trash(lst)
    stamp_added = any('deletedAt' not in x for x in lst)
    if dropped or stamp_added:
        save_trash(kept)
    return dropped

# ── 每月预算（budget.json：{ "2026-08": 3000.0, ... }）──
def load_budget():
    ensure_dir()
    if not os.path.exists(BUDGET_FILE):
        return {}
    try:
        with open(BUDGET_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def save_budget(b):
    ensure_dir()
    with open(BUDGET_FILE, 'w', encoding='utf-8') as f:
        json.dump(b, f, ensure_ascii=False, indent=2)

# ── 备份 ──
LEDGER_FILE_NAMES = ['records.json', 'trash.json', 'budget.json',
                     'recurring.json', 'settings.json']

# 标准备份目录名：20261018_160558
_BACKUP_NAME_RE = re.compile(r'^\d{8}_\d{6}$')

def do_backup():
    """备份 5 个账本文件到 data_backup/时间戳/，返回备份名"""
    ensure_dir()
    os.makedirs(BACKUP_DIR, exist_ok=True)
    name = time.strftime('%Y%m%d_%H%M%S')
    dest = os.path.join(BACKUP_DIR, name)
    os.makedirs(dest, exist_ok=True)
    for fn in LEDGER_FILE_NAMES:
        fp = os.path.join(DATA_DIR, fn)
        if os.path.isfile(fp):
            try:
                import shutil
                shutil.copy2(fp, os.path.join(dest, fn))
            except: pass
    prune_backups(5)
    return name

def prune_backups(keep=5):
    """只保留最新 keep 份正常备份；_pre_restore_ 临时备份保留最近 keep 份"""
    if not os.path.isdir(BACKUP_DIR): return
    import shutil
    # 正常备份：只留最新 keep 份
    # ⚠ 只算**标准时间戳目录**。
    #   2026-10-18 主人报「点完立即备份就没备份了」，真因就在这里：
    #   `upgrade_<stamp>`（升级时的**程序**备份，里面没有账本）
    #   也被收进来排序，而 'u'(117) > '2'(50) → upgrade_ 目录全排在前面，
    #   把 keep 名额全占了 → 真账本备份（包括刚生成的那份）全进 dirs[keep:] 被删。
    dirs = []
    for n in os.listdir(BACKUP_DIR):
        d = os.path.join(BACKUP_DIR, n)
        if os.path.isdir(d) and _BACKUP_NAME_RE.match(n):
            dirs.append(n)
    dirs.sort(reverse=True)  # 时间戳命名，字典序 = 时间序，最新在前
    for old in dirs[keep:]:
        try:
            shutil.rmtree(os.path.join(BACKUP_DIR, old), ignore_errors=True)
        except Exception:
            pass
    # upgrade_（程序备份）单独清理，不占账本备份的名额
    ups = [n for n in os.listdir(BACKUP_DIR)
           if os.path.isdir(os.path.join(BACKUP_DIR, n)) and n.startswith('upgrade_')]
    ups.sort(reverse=True)
    for old in ups[3:]:
        try:
            shutil.rmtree(os.path.join(BACKUP_DIR, old), ignore_errors=True)
        except Exception:
            pass
    temps = []
    for n in os.listdir(BACKUP_DIR):
        d = os.path.join(BACKUP_DIR, n)
        if os.path.isdir(d) and n.startswith('_pre_restore_'):
            temps.append(n)
    temps.sort(reverse=True)
    for old in temps[keep:]:
        try:
            shutil.rmtree(os.path.join(BACKUP_DIR, old), ignore_errors=True)
        except Exception:
            pass

def list_backups():
    """列出所有备份，带时间/记录数信息"""
    out = []
    if not os.path.isdir(BACKUP_DIR): return out
    for name in sorted(os.listdir(BACKUP_DIR), reverse=True):
        d = os.path.join(BACKUP_DIR, name)
        if not os.path.isdir(d): continue
        # _pre_restore_：恢复前的安全备份，不给用户点
        # upgrade_：升级时的**程序**备份（里面没有账本 json），不能当账本备份列出来
        if name.startswith('_pre_restore_') or name.startswith('upgrade_'):
            continue
        # 只认标准时间戳目录，避免把临时目录（_SAFE_/EMERGENCY_ 等）
        # 当备份列出来显示「0 条记录」
        if not _BACKUP_NAME_RE.match(name):
            continue
        recs = 0
        try:
            with open(os.path.join(d, 'records.json'), 'r', encoding='utf-8') as f:
                recs = len(json.load(f))
        except: pass
        mtime = os.path.getmtime(d)
        out.append({'name': name, 'time': time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(mtime)), 'records': recs})
    return out

# 账本文件白名单：备份/恢复只认这 5 个，别的一律不碰
LEDGER_FILES = ['records.json', 'trash.json', 'budget.json',
                'recurring.json', 'settings.json']


def restore_backup(name):
    """从备份恢复：先备份当前数据，再把备份内容拷回 data"""
    src = os.path.join(BACKUP_DIR, name)
    if not os.path.isdir(src):
        return False
    # ⚠ 必须确认这个目录里真的有账本！否则下面「先挪走、再拷回」的结果
    #   是账本被挪走后没东西拷回来 = data\ 直接变空。
    #   upgrade_* 这类只放程序的备份目录就属于这种，绝不能让它走到下一步。
    if not os.path.isfile(os.path.join(src, 'records.json')):
        return False
    # 先把当前数据挪进回收区，避免误删
    #   只挪 5 个账本文件：version.json / _check_records.json 这类
    #   非账本文件留在原地，否则恢复一次就被搬走了（旧版的暗伤）
    safe = os.path.join(BACKUP_DIR, '_pre_restore_' + time.strftime('%Y%m%d_%H%M%S'))
    os.makedirs(safe, exist_ok=True)
    for fn in LEDGER_FILES:
        fp = os.path.join(DATA_DIR, fn)
        if os.path.isfile(fp):
            try:
                import shutil
                shutil.move(fp, os.path.join(safe, fn))
            except: pass
    # 从备份拷回
    import shutil
    #   只拄 5 个账本文件，与上面「只挪 5 个」对称。
    #   旧写法把备份目录里所有文件全拄回去，
    #   老备份里的 version.json 会把版本号写回旧版（跟升级功能打架）。
    for fn in LEDGER_FILES:
        fp = os.path.join(src, fn)
        if os.path.isfile(fp):
            shutil.copy2(fp, os.path.join(DATA_DIR, fn))
    return True

# -- 外部备份包（导出 / 导入）--
# 2026-10-18 主人要求：能把备份存到 U 盘/别的文件夹，也能从那里选文件恢复。
# 备份包 = 一个 zip，里面平铺放 5 个账本 json（和 data_backup/<时间戳>/ 里的内容一致）。
BACKUP_PKG_MAX_BYTES = 100 * 1024 * 1024  # 100MB，账本 json 正常只有几百 KB


def export_backup_package(out_dir=None):
    """把当前 data 下的账本 json 打成一个 zip，返回 zip 路径（失败返回 ''）。

    只打包 5 个账本文件，不碰程序，也不带 _check_records.json 这类杂物。
    这样主人可以把它存到 U 盘/网盘，需要时再选回来。

    out_dir：输出目录。不传就放系统临时目录（导出下载用）；
             升级前自动备份会传 data_backup\\upgrade_<stamp>\\，
             这样带的就是同一个格式的「数据备份包」，以后能直接「从文件恢复」。
    """
    try:
        stamp = time.strftime('%Y%m%d_%H%M%S')
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
            zp = os.path.join(out_dir, '小爱记账备份_%s.zip' % stamp)
        else:
            zp = os.path.join(tempfile.gettempdir(), 'xiaoai_backup_%s.zip' % stamp)
        packed = 0
        with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as z:
            for fn in LEDGER_FILES:
                fp = os.path.join(DATA_DIR, fn)
                if os.path.isfile(fp):
                    z.write(fp, fn)
                    packed += 1
            readme = (
                '小爱记账 \u00b7 账本备份包\n'
                '======================================\n\n'
                '里面是 %d 个账本文件：\n'
                '  records.json   所有收支记录\n'
                '  trash.json     回收站\n'
                '  budget.json    预算设置\n'
                '  recurring.json 固定账单\n'
                '  settings.json  分类/图标设置\n\n'
                '备份时间：%s\n\n'
                '怎么恢复：\n'
                '1. 打开小爱记账 -> 设置 -> 数据备份\n'
                '2. 点「从文件恢复」，选中这个 zip 就行\n'
                '3. 恢复前程序会自动把当前数据另存一份，不怕弄错\n\n'
                '也可以直接解压，把里面的 json 覆盖到 小爱记账目录\\data\\。\n\n'
                '（升级前自动导出的备份包也是这个格式，恢复方法同上。）\n'
            ) % (packed, time.strftime('%Y-%m-%d %H:%M:%S'))
            z.writestr('备份说明.txt', readme.encode('utf-8'))
        return zp
    except Exception:
        return ''


def _safe_backup_name(name):
    """只取 basename，挡掉 ../ 之类的路径穿越；不在白名单就返回空串。"""
    base = os.path.basename((name or '').replace('\\', '/'))
    return base if base in LEDGER_FILES else ''


def import_backup_package(raw):
    """从上传的 zip 恢复账本。返回 (ok, msg)。

    安全点：
      1. 只认 5 个白名单 json，别的文件一律忽略（程序文件绝不动）；
      2. zip 内路径只取 basename，砍掉任何 ../ 前缀；
      3. 解压总大小先卡 100MB，防 zip 炸弹；
      4. 恢复前把当前 data 全量挪到 data_backup/_pre_restore_<时间戳>/；
      5. 先校验每个 json 能解析，解析不过直接返回，不留下半截数据。
    """
    if not raw:
        return False, '备份文件是空的呢'
    if len(raw) > BACKUP_PKG_MAX_BYTES:
        return False, '备份包太大（超过 100MB），不像账本备份'
    tmp = os.path.join(tempfile.gettempdir(), '_xiaoai_rst_%d.zip' % os.getpid())
    staged = {}
    try:
        with open(tmp, 'wb') as f:
            f.write(raw)
        if not zipfile.is_zipfile(tmp):
            return False, '这不是一个 zip 备份包呢'
        zf = zipfile.ZipFile(tmp, 'r')
        # 同名文件多层时优先取层级最浅的（根目录那个）
        infos = sorted([i for i in zf.infolist() if not i.filename.endswith('/')],
                       key=lambda i: i.filename.count('/'))
        total = 0
        for i in infos:
            fn = _safe_backup_name(i.filename)
            if not fn:
                continue
            total += max(0, i.file_size)
            if total > BACKUP_PKG_MAX_BYTES:
                return False, '备份包解开后太大了，不像正常账本备份'
            if fn not in staged:
                staged[fn] = zf.read(i)
        if 'records.json' not in staged:
            return False, '这个包里没有 records.json，不像账本备份'
        # 先校验 json 能不能解析，坏了就一个字节都不动
        for fn, blob in staged.items():
            try:
                json.loads(blob.decode('utf-8'))
            except Exception:
                return False, '%s 内容损坏，没法恢复' % fn
        # 先把当前账本整体挪出去（和 restore_backup 同一套安全网）
        #   只动 5 个账本文件：version.json 这类程序状态必须原地不动，
        #   否则恢复一次版本号就丢了。
        ensure_dir()
        safe = os.path.join(BACKUP_DIR, '_pre_restore_' + time.strftime('%Y%m%d_%H%M%S'))
        os.makedirs(safe, exist_ok=True)
        for fn in LEDGER_FILES:
            fp = os.path.join(DATA_DIR, fn)
            if os.path.isfile(fp):
                try:
                    shutil.move(fp, os.path.join(safe, fn))
                except Exception:
                    pass
        for fn, blob in staged.items():
            with open(os.path.join(DATA_DIR, fn), 'wb') as f:
                f.write(blob)
        prune_backups(5)
        return True, '已恢复 %d 个数据文件' % len(staged)
    except Exception as e:
        return False, '恢复失败：%s' % e
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


def delete_backup(name):
    import shutil
    d = os.path.join(BACKUP_DIR, name)
    if os.path.isdir(d):
        shutil.rmtree(d, ignore_errors=True)
        return True
    return False

# ── 固定/周期性账单（recurring.json）──
# 结构：[{ id, name, type, category, amount, note, freq, day, month, start,
#           lastAuto, end, enabled, createdAt }]
# freq: monthly(按 day 每月) / yearly(按 month+day 每年)
# lastAuto: 最近一次自动生成的日期 'YYYY-MM-DD'，用于去重
SYS_RECUR_WINDOW_MONTHS = 3  # 最多补最近 3 期，避免积压记录一次性爆炸
SYS_RECUR_MAX_AGE_DAYS = 400  # 超过一年的旧期数不再补

def load_recurring():
    ensure_dir()
    if not os.path.exists(RECURRING_FILE):
        return []
    try:
        with open(RECURRING_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []

def save_recurring(items):
    ensure_dir()
    with open(RECURRING_FILE, 'w', encoding='utf-8') as f:
        json.dump(items, f, ensure_ascii=False, indent=2)

def parse_ymd(s):
    """'YYYY-MM-DD' -> (y, m, d)，失败返回 None"""
    import datetime as _dt
    try:
        parts = str(s).split('-')
        return _dt.date(int(parts[0]), int(parts[1]), int(parts[2]))
    except Exception:
        return None

def day_clamp(y, m, d):
    """把 day 夹到该月合法范围内（31 号在小月自动落到月末）"""
    import calendar
    try:
        last = calendar.monthrange(y, m)[1]
    except Exception:
        return None
    return min(max(int(d), 1), last)

def recur_occ_dates(item, today=None, window_months=SYS_RECUR_WINDOW_MONTHS):
    """返回本期应自动生成的日期列表（字符串）
    规则：从 start 到 today 之间所有到期日，最多取最近 window_months 期，
    且只保留最近 SYS_RECUR_MAX_AGE_DAYS 天内的（更早的不补历史），
    已生成过的（<= lastAuto）不重复。
    """
    import datetime as _dt
    if today is None:
        today = _dt.date.today()
    freq = item.get('freq', 'monthly')
    start = parse_ymd(item.get('start') or '')
    if start is None:
        start = today
    end = parse_ymd(item.get('end') or '')
    last = item.get('lastAuto') or ''
    raw = []
    if freq == 'yearly':
        mo = int(item.get('month') or start.month)
        dy = int(item.get('day') or start.day)
        for y in range(start.year, today.year + 2):
            d = day_clamp(y, mo, dy)
            if d is None:
                continue
            raw.append(_dt.date(y, mo, d))
    else:
        dy = int(item.get('day') or start.day)
        y, m = start.year, start.month
        guard = 0
        while guard < 600:
            guard += 1
            d = day_clamp(y, m, dy)
            if d is not None:
                raw.append(_dt.date(y, m, d))
            m += 1
            if m > 12:
                m = 1
                y += 1
            if _dt.date(y, m, 1) > today + _dt.timedelta(days=40):
                break
    out = []
    for cur in raw:
        if cur < start or cur > today:
            continue
        if end is not None and cur > end:
            continue
        s = cur.strftime('%Y-%m-%d')
        if last and s <= last:
            continue
        if (today - cur).days > SYS_RECUR_MAX_AGE_DAYS:
            continue
        out.append((cur, s))
    out.sort(key=lambda x: x[0])
    keep = int(window_months)
    if keep > 0:
        out = out[-keep:]
    return [s for _, s in out]

def maybe_apply_recurring():
    """自动到期入账：按固定账单生成缺失的记录，返回生成条数"""
    items = load_recurring()
    if not items:
        return 0
    recs = load_records()
    nid = next_id(recs)
    now = int(time.time() * 1000)
    added = 0
    changed = False
    for it in items:
        if it.get('enabled') is False:
            continue
        amt = float(it.get('amount') or 0)
        if amt <= 0:
            continue
        item_changed = False
        for ds in recur_occ_dates(it):
            dup = False
            for r in recs:
                if r.get('recurringId') == it.get('id') and r.get('date') == ds:
                    dup = True
                    break
            if not dup:
                recs.append({
                    'id': nid,
                    'date': ds,
                    'type': it.get('type', 'expense'),
                    'category': it.get('category', '其他杂项'),
                    'sub': it.get('sub', '') or '',
                    'amount': round(amt, 2),
                    'note': it.get('note', '') or it.get('name', ''),
                    'createdAt': now + nid,
                    'recurringId': it.get('id'),
                    'auto': True,
                })
                added += 1
            if not it.get('lastAuto') or ds > it.get('lastAuto'):
                it['lastAuto'] = ds
                item_changed = True
            nid += 1
        if item_changed:
            changed = True
    if added:
        save_records(recs)
    if changed:
        save_recurring(items)
    return added

def next_id(records):
    return (max(r.get('id',0) for r in records) + 1) if records else 1

class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_cors()
        self.end_headers()

    def send_cors(self):
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Access-Control-Allow-Methods','GET,POST,DELETE,PUT,OPTIONS')
        self.send_header('Access-Control-Allow-Headers','Content-Type')

    def json_resp(self, data, status=200):
        self.send_response(status)
        self.send_cors()
        self.send_header('Content-Type','application/json;charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def read_body(self):
        length = int(self.headers.get('Content-Length',0))
        if not length: return {}
        return json.loads(self.rfile.read(length).decode('utf-8'))

    def import_xlsx(self, body):
        if not HAS_XL:
            self.json_resp({'error':'openpyxl not available'},500); return
        import io, base64, re
        import datetime as _dt
        b64 = body.get('file','')
        if not b64: self.json_resp({'error':'No file data'},400); return
        try:
            buf = io.BytesIO(base64.b64decode(b64))
            wb = openpyxl.load_workbook(buf)
            r = load_records()
            current_id = next_id(r)
            now = int(time.time()*1000)
            imported = 0
            errors = []

            # 支出/收入分类集合（用于无 Type 列时智能判断收支）
            EXPENSE_CATS_SET = set(DEFAULT_EXPENSE + ['早餐','午餐','晚餐','水果零食','零食水果','甜点饮品','餐饮美食','早午晚餐','租房水电','行车交通','医疗保健','衣服饰品','休闲娱乐','游戏娱乐','电子数码','日用百货','交流通讯','学习充值','人情社交','宠物花销','理财存入','理财基金','理财亏损','存入存款'])
            INCOME_CATS_SET = set(DEFAULT_INCOME + ['奖金','兼职','理财取出','理财收入','取出存款','红包','转账'])

            # 表头列名 → 语义映射
            def norm(s):
                return str(s).strip().lower().replace(' ','').replace('_','')
            def col_sem(name):
                n = norm(name)
                if n in ('date','日期','时间','day'): return 'date'
                if n in ('type','类型','收支','收支类型','收入支出','类别类型'): return 'type'
                if n in ('category','分类','类别','类目','cat','支出类别','收入类别'): return 'category'
                if n in ('二级分类','子分类','二级','小类','subcategory','subcat','sub'): return 'subcategory'
                if n in ('amount','金额','money','钱','费用','花费','支出金额','收入金额'): return 'amount'
                if n in ('note','备注','说明','描述','notes','memo'): return 'note'
                return None

            # 清洗金额：去 ¥ , 空格，支持负数
            def clean_amount(v):
                if v is None or v == '': return None
                if isinstance(v, (int, float)):
                    return float(v)
                s = re.sub(r'[¥￥,，\s元]', '', str(v).strip())
                if s == '' or s in ('-','+'): return None
                try: return float(s)
                except ValueError: return None

            # 日期标准化为 YYYY-MM-DD
            def clean_date(v):
                if v is None or v == '': return ''
                if isinstance(v, _dt.datetime): return v.strftime('%Y-%m-%d')
                if isinstance(v, _dt.date): return v.strftime('%Y-%m-%d')
                s = str(v).strip()
                # Excel 序列号（数字）
                if re.fullmatch(r'\d{5}', s):
                    try:
                        return (_dt.date(1899,12,30) + _dt.timedelta(days=int(s))).strftime('%Y-%m-%d')
                    except Exception: return ''
                m = re.match(r'(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?', s)
                if m:
                    return '%s-%02d-%02d' % (m.group(1), int(m.group(2)), int(m.group(3)))
                return s

            def make_rec(date, typ, cat, amt, note, sub=''):
                nonlocal current_id, imported
                rec = {'id':current_id,'date':date,'type':typ,'category':cat,'amount':round(amt,2),'note':note,'createdAt':now+current_id}
                if sub: rec['sub'] = sub
                r.append(rec)
                current_id += 1
                imported += 1

            def guess_type(cat):
                # 无 Type 列时按分类名判断：收入类目→income，其余→expense
                if cat in INCOME_CATS_SET: return 'income'
                return 'expense'

            # 解析一个 sheet：cols 是 {语义: 列索引}，header_row 是表头行号（数据从 header_row+1 开始）
            def parse_sheet(ws, cols, header_row, force_type=None):
                nonlocal imported
                for row in ws.iter_rows(min_row=header_row+1, values_only=True):
                    if not row or not any(row): continue
                    date = clean_date(row[cols['date']]) if 'date' in cols else ''
                    if not date: continue
                    cat = str(row[cols['category']]).strip() if 'category' in cols and row[cols['category']] else '其他'
                    _sub = ''
                    if 'subcategory' in cols:
                        _si = cols['subcategory']
                        if len(row) > _si and row[_si] is not None:
                            _sub = str(row[_si]).strip()
                    amt = clean_amount(row[cols['amount']]) if 'amount' in cols and row[cols['amount']] is not None else None
                    if amt is None: continue
                    note = str(row[cols['note']]).strip() if 'note' in cols and len(row)>cols['note'] and row[cols['note']] else ''
                    if force_type:
                        typ = force_type
                    elif 'type' in cols and row[cols['type']] is not None:
                        ts = str(row[cols['type']]).strip()
                        typ = 'expense' if ('支' in ts or 'expense' in ts.lower()) else ('income' if ('收' in ts or 'income' in ts.lower()) else guess_type(cat))
                    else:
                        typ = guess_type(cat)
                    make_rec(date, typ, cat, amt, note, _sub)

            # 找每个 sheet 的表头行和列语义
            def detect_sheet(ws):
                for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
                    if not row: continue
                    cols = {}
                    for ci, cell in enumerate(row):
                        sem = col_sem(cell) if cell is not None else None
                        if sem and sem not in cols:
                            cols[sem] = ci
                    if 'date' in cols and 'amount' in cols:
                        return ri, cols
                return None, None

            snames = [s.title for s in wb.worksheets]
            sheets_done = set()
            if '支出' in snames:
                ws = wb['支出']; hr, cols = detect_sheet(ws)
                if hr: parse_sheet(ws, cols, hr, force_type='expense'); sheets_done.add('支出')
            if '收入' in snames:
                ws = wb['收入']; hr, cols = detect_sheet(ws)
                if hr: parse_sheet(ws, cols, hr, force_type='income'); sheets_done.add('收入')
            # 其他 sheet（含单 sheet 文件）：自动检测表头解析
            for ws in wb.worksheets:
                if ws.title in sheets_done: continue
                hr, cols = detect_sheet(ws)
                if hr: parse_sheet(ws, cols, hr)

            save_records(r)
            self.json_resp({'ok':True,'imported':imported,'errors':errors[:10]})
        except Exception as e:
            self.json_resp({'error':str(e)},500)

    def serve_static(self, path):
        try:
            if path == '/' or path == '':
                path = '/记账本-v2.html'
            # 优先从 exe 同目录加载（方便热更新 html/css/js）
            unq = urllib.parse.unquote(path)
            fpath = os.path.join(EXE_DIR, unq.lstrip('/'))
            if not os.path.isfile(fpath):
                fpath = BUNDLE_DIR + urllib.parse.unquote(path)
            if not os.path.isfile(fpath):
                self.send_response(404)
                self.send_cors()
                self.end_headers()
                return
            ext = os.path.splitext(fpath)[1].lower()
            ct = MIME.get(ext, 'application/octet-stream')
            self.send_response(200)
            self.send_cors()
            self.send_header('Content-Type', ct)
            self.end_headers()
            with open(fpath, 'rb') as f:
                self.wfile.write(f.read())
        except Exception as e:
            try:
                self.send_response(500); self.send_cors(); self.send_header('Content-Type','text/plain'); self.end_headers()
                self.wfile.write(str(e).encode('utf-8'))
            except: pass

    # ── 导出完整项目包（干净空白账本，可直接分享给别人）──
    # 只放「程序本体 + 空数据」，绝不带主人真实数据。
    # ── 一键升级：只换程序文件，数据一字不动 ──
    def upgrade_from_zip(self, body):
        """接收分享 ZIP -> 校验 -> 备份旧程序 -> 解压程序文件 -> 写版本号

        ⚠️ 绝不解压到 data\：账本、预算、设置全部保留。
        ⚠️ exe 正在跑不能直接覆盖 -> 先存成 .new，重启时换上去。
        """
        import base64
        b64 = body.get('file', '')
        if not b64:
            self.json_resp({'error': '没有收到文件'}, 400)
            return
        try:
            raw = base64.b64decode(b64)
        except Exception:
            self.json_resp({'error': '文件内容损坏，无法解码'}, 400)
            return
        if not raw:
            self.json_resp({'error': '文件是空的'}, 400)
            return
        if len(raw) > 200 * 1024 * 1024:
            self.json_resp({'error': '文件太大（超过 200MB）'}, 400)
            return

        tmp = os.path.join(tempfile.gettempdir(), '_xiaoai_up_%d.zip' % os.getpid())
        try:
            with open(tmp, 'wb') as f:
                f.write(raw)
            if not zipfile.is_zipfile(tmp):
                self.json_resp({'error': '这不是一个 zip 压缩包呢'}, 400)
                return

            zf = zipfile.ZipFile(tmp, 'r')
            try:
                infos = [i for i in zf.infolist() if not i.filename.endswith('/')]

                # —— 防 zip 炸弹：解压后总大小先卡住 ——
                total_unpacked = sum(max(0, i.file_size) for i in infos)
                if total_unpacked > 600 * 1024 * 1024:
                    self.json_resp({'error': '这个压缩包解开后有 %d MB，太大了，'
                                             '不像正常项目包呢'
                                             % (total_unpacked // (1024 * 1024))}, 400)
                    return

                # —— 同名文件在多层级出现时，优先取根目录那个 ——
                best = {}
                for i in infos:
                    bn = i.filename.replace('\\', '/').split('/')[-1]
                    depth = i.filename.replace('\\', '/').count('/')
                    if bn not in best or depth < best[bn][0]:
                        best[bn] = (depth, i)

                # —— 读新版本号（不解压到磁盘）——
                new_ver = ''
                vi = best.get('version.json')
                if vi:
                    try:
                        info = json.loads(zf.read(vi[1]).decode('utf-8'))
                        if isinstance(info, dict) and info.get('version'):
                            new_ver = str(info['version']).strip()[:32]
                    except Exception:
                        pass

                # —— 校验：必须含本项目的程序文件 ——
                bases = set(best.keys())

                # ── 检查包里有没有带文档 ──
                #   包内缺文档时升级不会报错，但《更新日志》《使用手册》会一直停在旧版，
                #   主人会以为「升级没生效」。这里记下来，结果里明确告知。
                docs_missing = [fn for fn in ('更新日志.md', '使用手册.md')
                                if fn not in bases]

                program = [fn for fn in UPGRADE_PROGRAM_FILES if fn in bases]
                if not program:
                    self.json_resp({'error': '这个压缩包里没有小爱记账的程序文件，'
                                             '可能选错包了呢'}, 400)
                    return

                cur = read_version_file()
                cur_ver = cur.get('version', APP_VERSION)

                # —— 备份旧程序（出事可以换回去）——
                stamp = time.strftime('%Y%m%d_%H%M%S')
                bak_dir = os.path.join(BACKUP_DIR, 'upgrade_' + stamp)
                try:
                    os.makedirs(bak_dir, exist_ok=True)
                    for fn in program:
                        fp = os.path.join(EXE_DIR, fn)
                        if os.path.isfile(fp):
                            shutil.copy2(fp, os.path.join(bak_dir, fn))
                except Exception as ex:
                    self.json_resp({'error': '备份旧程序失败：%s' % ex}, 500)
                    return

                # —— 升级前自动导出一份账本备份（2026-10-18 主人要求）——
                #   放在进来就做，不管后面解压成不成功，备份都已经在了。
                ledger_bak = backup_ledger_before_upgrade(bak_dir)

                # —— 解压程序文件（data\ 一律跳过）——
                updated, staged, skipped = [], [], []
                restart_needed = False
                for base_n in program:
                    zi = best.get(base_n)
                    if not zi:
                        continue
                    if zi[1].file_size > 300 * 1024 * 1024:
                        skipped.append('%s（单个文件 %d MB，太大了）'
                                       % (base_n, zi[1].file_size // (1024 * 1024)))
                        continue
                    try:
                        blob = zf.read(zi[1])
                    except Exception as ex:
                        skipped.append('%s（%s）' % (base_n, ex))
                        continue
                    # —— 内容有效性：防止选错包把程序写成垃圾 ——
                    bad = _program_blob_problem(base_n, blob)
                    if bad:
                        skipped.append('%s（%s）' % (base_n, bad))
                        continue
                    if len(blob) > 300 * 1024 * 1024:
                        skipped.append('%s（解压后太大）' % base_n)
                        continue
                    dest = os.path.join(EXE_DIR, base_n)

                    # exe 正在运行，Windows 不让覆盖 -> 先存 .new
                    if base_n.lower().endswith('.exe'):
                        ok, why = _write_staged_exe(dest, blob)
                        if ok:
                            staged.append(base_n)
                            restart_needed = True
                        else:
                            skipped.append('%s（%s）' % (base_n, why))
                        continue

                    # 先写临时文件再原子替换，避免读到一半
                    tmpp = dest + '.updatetmp'
                    try:
                        with open(tmpp, 'wb') as f:
                            f.write(blob)
                        try:
                            os.replace(tmpp, dest)
                        except Exception:
                            if os.path.isfile(dest):
                                os.remove(dest)
                            os.replace(tmpp, dest)
                        updated.append(base_n)
                    except Exception as ex:
                        skipped.append('%s（%s）' % (base_n, ex))
                        try:
                            if os.path.isfile(tmpp):
                                os.remove(tmpp)
                        except Exception:
                            pass

                # —— 写版本号 ——
                to_ver = new_ver or cur_ver
                is_downgrade = False
                try:
                    is_downgrade = compare_version(to_ver, cur_ver) < 0
                except Exception:
                    is_downgrade = False
                write_version_file(to_ver,
                                   '从 %s 升级到 %s' % (cur_ver, to_ver))

                # —— 需要重启：只写 bat，**绝不提前拉起助手** ——
                # ⚠⚠ 2026-10-18 重大修正：
                #   以前这里直接 _spawn_upgrade_helper()，助手一被拉起就在后台
                #   盯着 `.new`；一旦旧进程退了（或被杀），它就自己启动程序
                #   → 程序每次启动都自动开浏览器 → 主人没点重启也会弹窗，
                #   而且反复重启就弹二三十次。
                #   现在：升级只负责把 `.new` 写好 + 写一个 bat，
                #   **什么时候重启完全由主人点按钮决定**。
                helper = ''
                helper_ready = False
                if restart_needed:
                    try:
                        helper = os.path.join(EXE_DIR, '_upgrade_restart.bat')
                        with open(helper, 'w', encoding='utf-8', newline='') as f:
                            f.write(_upgrade_helper_bat())
                        helper_ready = True   # bat 已就位，点按钮就能重启
                    except Exception:
                        helper = ''

                self.json_resp({
                    'ok': True,
                    'fromVersion': cur_ver,
                    'toVersion': to_ver,
                    'downgrade': is_downgrade,
                    'updated': updated,
                    'staged': staged,
                    'skipped': skipped,
                    'docsMissing': docs_missing,
                    'backupDir': bak_dir,
                    'ledgerBackup': ledger_bak,
                    'restartNeeded': restart_needed,
                    'helper': helper,
                    'helperReady': helper_ready,
                })
            finally:
                try:
                    zf.close()
                except Exception:
                    pass
        except Exception as ex:
            try:
                self.json_resp({'error': '升级失败：%s' % ex}, 500)
            except Exception:
                pass
        finally:
            try:
                os.remove(tmp)
            except Exception:
                pass

    def send_backup_zip(self):
        """把当前账本打成一个 zip 给浏览器下载（存 U 盘/网盘用）。"""
        from urllib.parse import quote
        zp = export_backup_package()
        if not zp or not os.path.isfile(zp):
            self.json_resp({'error': '打包备份失败'}, 500)
            return
        try:
            size = os.path.getsize(zp)
            stamp = time.strftime('%Y%m%d_%H%M%S')
            fname = '小爱记账备份_%s.zip' % stamp
            self.send_response(200)
            self.send_cors()
            self.send_header('Content-Type', 'application/zip')
            self.send_header('Content-Length', str(size))
            self.send_header('Content-Disposition',
                             "attachment; filename=catbook_backup_%s.zip; filename*=UTF-8''%s"
                             % (stamp, quote(fname)))
            self.end_headers()
            with open(zp, 'rb') as f:
                while True:
                    chunk = f.read(262144)
                    if not chunk:
                        break
                    self.wfile.write(chunk)
        finally:
            try:
                os.remove(zp)
            except Exception:
                pass

    def send_project_zip(self):
        import datetime as _dt
        from urllib.parse import quote
        tmp = os.path.join(tempfile.gettempdir(), '_xiaoai_proj_{}.zip'.format(os.getpid()))
        try:
            # 用「升级白名单」做源头，保证导出的包天然就是合格的升级包。
            #   2026-10-18 主人报「升级后更新日志/使用手册没更新」——
            #   这里以前只写死 5 个文件，漏了 使用手册.md / 更新日志.md，
            #   升级时包里没它们，于是文档永远停在旧版。
            #   使用说明.txt / 启动.bat 由下面 writestr 现生成，这里排除掉防重名。
            names = [n for n in UPGRADE_PROGRAM_FILES
                     if n not in ('使用说明.txt', '启动.bat')]
            picked, seen = [], set()
            for fn in names:
                for base in (EXE_DIR, BUNDLE_DIR):
                    fp = os.path.join(base, fn)
                    if os.path.isfile(fp) and fp not in seen:
                        picked.append((fn, fp))
                        seen.add(fp)
                        break
            # 空白 data：分类用默认表，不带主人自定义过的分类/二级
            empty_settings = {
                'expenseCats': list(DEFAULT_EXPENSE),
                'incomeCats': list(DEFAULT_INCOME),
                'customIcons': {},
                'subCats': {}
            }
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
                for fn, fp in picked:
                    z.write(fp, fn)
                z.writestr('data/records.json', '[]')
                z.writestr('data/trash.json', '[]')
                z.writestr('data/budget.json', '{}')
                z.writestr('data/recurring.json', '[]')
                z.writestr('data/settings.json',
                           json.dumps(empty_settings, ensure_ascii=False, indent=2))
                # 带上版本号 + 清单，方便对方「一键升级」时比对
                z.writestr('version.json', json.dumps(
                    {'version': APP_VERSION,
                     'programFiles': UPGRADE_PROGRAM_FILES},
                    ensure_ascii=False, indent=2))
                z.writestr('data/version.json', json.dumps(
                    {'version': APP_VERSION,
                     'updatedAt': time.strftime('%Y-%m-%d %H:%M:%S'),
                     'note': '分享包初始版本'},
                    ensure_ascii=False, indent=2))
                z.writestr('使用说明.txt', _project_readme())
                z.writestr('启动.bat', _project_launcher())
            size = os.path.getsize(tmp)
            stamp = _dt.datetime.now().strftime('%Y%m%d')
            # 文件名带版本号（2026-10-18 主人要求）：小爱记账_空白账本_v2.15.9.zip
            #   原来的日期那一栏换成版本号，一眼看出是哪个版本打的包。
            fname = '小爱记账_空白账本_v{}.zip'.format(APP_VERSION)
            self.send_response(200)
            self.send_cors()
            self.send_header('Content-Type', 'application/zip')
            self.send_header('Content-Length', str(size))
            self.send_header('Content-Disposition',
                             "attachment; filename=catbook_v{}.zip; filename*=UTF-8''{}"
                             .format(APP_VERSION, quote(fname)))
            self.end_headers()
            with open(tmp, 'rb') as f:
                while True:
                    chunk = f.read(262144)
                    if not chunk:
                        break
                    self.wfile.write(chunk)
        except Exception as e:
            try:
                self.json_resp({'error': str(e)}, 500)
            except Exception:
                pass
        finally:
            try:
                os.remove(tmp)
            except Exception:
                pass

    def send_xlsx(self, from_ym=None, to_ym=None):
        if not HAS_XL:
            self.send_response(500); self.send_cors(); self.send_header('Content-Type','text/plain'); self.end_headers()
            self.wfile.write(b'openpyxl not available'); return
        if from_ym and to_ym and from_ym > to_ym:
            from_ym, to_ym = to_ym, from_ym
        tmp = os.path.join(tempfile.gettempdir(), '_xiaoai_export_{}.xlsx'.format(os.getpid()))
        wb = openpyxl.Workbook()
        ws_exp = wb.active
        ws_exp.title = '支出'
        ws_inc = wb.create_sheet('收入')
        ws_exp.append(['Date','Category','Subcategory','Amount','Note'])
        ws_inc.append(['Date','Category','Subcategory','Amount','Note'])
        for r in load_records():
            d = r.get('date','') or ''
            ym = d[:7]
            if from_ym and ym and ym < from_ym: continue
            if to_ym and ym and ym > to_ym: continue
            if from_ym and not ym: continue
            row = [d, r.get('category',''), r.get('sub',''), r.get('amount',0), r.get('note','')]
            if r.get('type')=='expense': ws_exp.append(row)
            else: ws_inc.append(row)
        wb.save(tmp)
        wb.close()
        self.send_response(200)
        self.send_cors()
        self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.send_header('Content-Disposition','attachment; filename="records.xlsx"')
        self.end_headers()
        with open(tmp,'rb') as f:
            self.wfile.write(f.read())
        try: os.remove(tmp)
        except: pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        p = parsed.path; q = urllib.parse.parse_qs(parsed.query)
        try:
            if p == '/api/settings':
                self.json_resp(load_settings())
            elif p == '/api/records':
                r = load_records()
                month = q.get('month',[None])[0]; stype = q.get('type',[None])[0]
                search = q.get('search',[None])[0]; sort = q.get('sort',['newest'])[0]
                limit = q.get('limit',[None])[0]
                if month: r = [x for x in r if x.get('date','').startswith(month)]
                if stype: r = [x for x in r if x.get('type')==stype]
                if search:
                    s = search.lower(); r = [x for x in r if s in x.get('category','').lower() or s in x.get('note','').lower()]
                r.sort(key=lambda x: (x.get('date',''), x.get('createdAt',0)), reverse=(sort!='oldest'))
                if limit: r = r[:int(limit)]
                self.json_resp(r)
            elif p == '/api/trash':
                # 惰性清理：按保留天数先清掉过期的，再把剩下的返回
                dropped = auto_purge_trash()
                self.json_resp({'records': load_trash(), 'purged': dropped,
                                'keepDays': trash_keep_days()})
            elif p == '/api/budget':
                self.json_resp(load_budget())
            elif p == '/api/recurring':
                self.json_resp(load_recurring())
            elif p == '/api/recurring/apply':
                added = maybe_apply_recurring()
                self.json_resp({'ok':True,'added':added,'items':load_recurring()})
            elif p == '/api/backups':
                self.json_resp({'backups': list_backups()})
            elif p == '/api/backup/export':
                self.send_backup_zip()
            elif p == '/api/export':
                self.json_resp({'version':2,'exportedAt':time.strftime('%Y-%m-%dT%H:%M:%S.000Z',time.gmtime()),'records':load_records(),'settings':load_settings()})
            elif p == '/api/export_xlsx':
                self.send_xlsx(q.get('from',[None])[0], q.get('to',[None])[0])
            elif p == '/api/version':
                v = read_version_file()
                file_ver = v.get('version', APP_VERSION)
                # runningVersion = 当前这个进程自带的版本
                # pendingVersion = data/version.json 里的版本（升级后需重启才生效）
                pending = file_ver if compare_version(file_ver, APP_VERSION) != 0 else ''
                self.json_resp({'appVersion': APP_VERSION,
                                'version': file_ver,
                                'runningVersion': APP_VERSION,
                                'pendingVersion': pending,
                                'updatedAt': v.get('updatedAt', ''),
                                'note': v.get('note', ''),
                                'programFiles': UPGRADE_PROGRAM_FILES})
            elif p == '/api/export_project':
                self.send_project_zip()
            elif p == '/api/shutdown':
                _q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
                _relaunch = (_q.get('relaunch', ['0'])[0] or '0').lower() in ('1', 'true', 'yes')
                # ⚠ 2026-10-18：relaunch=1 只会在**主人点了「立即重启」**时出现，
                #   这是唯一允许拉起助手的地方（升级本身绝不拉起）。
                #   如果真活不到就自己开一个 \"_restart_once.bat\"？不需要：
                #   前端会在 60 秒后给兜底提示。
                if _relaunch:
                    _hpath = os.path.join(EXE_DIR, '_upgrade_restart.bat')
                    if os.path.isfile(_hpath):
                        # 先清掉上一轮的 launched 标记，允许这一次启动
                        try:
                            _dm = os.path.join(EXE_DIR, '_helper.launched')
                            if os.path.isfile(_dm):
                                os.remove(_dm)
                        except Exception:
                            pass
                        _spawn_upgrade_helper(_hpath)
                self.json_resp({'ok': True, 'relaunch': _relaunch,
                                'message': '服务器关闭喵~'})
                # 1) 切走工作目录释放锁 2) 发送退出到全部客户端 3) 等响应写完再退出
                def _force_exit():
                    import time
                    try: os.chdir(os.environ['TEMP'])
                    except: pass
                    time.sleep(0.3)
                    os._exit(0)
                threading.Thread(target=_force_exit, daemon=True).start()
            elif p == '/ping':
                self.json_resp({'ok':True})
            else:
                self.serve_static(p)
        except Exception as e:
            try: self.send_response(500); self.send_cors(); self.send_header('Content-Type','text/plain'); self.end_headers(); self.wfile.write(str(e).encode('utf-8'))
            except Exception:
                pass

    def do_POST(self):
        p = urllib.parse.urlparse(self.path).path; body = self.read_body()
        try:
            if p == '/api/records':
                r = load_records(); new_id = next_id(r)
                rec = {'id':new_id,'date':body.get('date',time.strftime('%Y-%m-%d')),'type':body.get('type','expense'),'category':body.get('category','其他'),'sub':body.get('sub',''),'amount':round(float(body.get('amount',0)),2),'note':body.get('note',''),'createdAt':int(time.time()*1000)+new_id}
                # 退款关联：创建退款时若指定 refundOf，同时写原支出的 refundedBy
                ro = body.get('refundOf')
                if ro is not None and rec.get('category') == REFUND_CAT:
                    try: ro = int(ro)
                    except (TypeError, ValueError): ro = None
                    if ro is not None:
                        target = None
                        for rr in r:
                            if rr.get('id') == ro and rr.get('type') == 'expense':
                                target = rr; break
                        if target is not None:
                            rec['refundOf'] = ro
                            target['refundedBy'] = new_id
                r.append(rec); save_records(r); self.json_resp(rec,201)
            elif p == '/api/settings':
                s = load_settings(); s.update(body); save_settings(s); self.json_resp({'ok':True})
            elif p == '/api/budget':
                b = load_budget()
                month = body.get('month','')
                amount = body.get('amount')
                # 支持存 reset 列表：body 带 list 字段时更新 BUDGETS['reset']
                if 'list' in body:
                    if body.get('list'):
                        b['reset'] = [str(x) for x in body.get('list')]
                    else:
                        b.pop('reset', None)
                    save_budget(b)
                    self.json_resp(b); return
                # 支持存「关闭预算」列表 off + 关闭起点 offFrom
                # off  = 精确关掉的月份 ['2026-11', ...]
                # offFrom = 关闭默认预算后，从该月起往后都没有预算
                if 'off' in body or 'offFrom' in body:
                    if 'off' in body:
                        lst = body.get('off') or []
                        if lst:
                            b['off'] = [str(x) for x in lst]
                        else:
                            b.pop('off', None)
                    if 'offFrom' in body:
                        f = body.get('offFrom')
                        if f:
                            b['offFrom'] = str(f)
                        else:
                            b.pop('offFrom', None)
                    save_budget(b)
                    self.json_resp(b); return
                if not month:
                    self.json_resp({'error':'缺少月份'},400); return
                # 金额为空/None/<=0 时删除该月预算
                if amount is None or amount == '':
                    b.pop(month, None)
                else:
                    try:
                        amt = float(amount)
                    except (TypeError, ValueError):
                        self.json_resp({'error':'金额格式不对'},400); return
                    if amt <= 0:
                        b.pop(month, None)
                    else:
                        b[month] = round(amt, 2)
                save_budget(b)
                self.json_resp(b)
            elif p == '/api/recurring':
                items = load_recurring()
                nid = (max([int(x.get('id') or 0) for x in items]) + 1) if items else 1
                name = str(body.get('name','')).strip()
                amt = body.get('amount')
                try:
                    amt = float(amt)
                except (TypeError, ValueError):
                    self.json_resp({'error':'金额格式不对'},400); return
                if amt <= 0:
                    self.json_resp({'error':'金额必须大于0'},400); return
                if not name:
                    name = body.get('category','固定支出')
                freq = 'yearly' if body.get('freq') == 'yearly' else 'monthly'
                def _int(v, dflt):
                    try: return int(v)
                    except (TypeError, ValueError): return dflt
                start = str(body.get('start') or time.strftime('%Y-%m-%d'))
                it = {
                    'id': nid,
                    'name': name,
                    'type': 'income' if body.get('type') == 'income' else 'expense',
                    'category': body.get('category','其他杂项'),
                    'sub': (body.get('sub') or '').strip(),
                    'amount': round(amt, 2),
                    'note': body.get('note',''),
                    'freq': freq,
                    'day': _int(body.get('day'), parse_ymd(start).day if parse_ymd(start) else 1),
                    'month': _int(body.get('month'), parse_ymd(start).month if parse_ymd(start) else 1),
                    'start': start,
                    'lastAuto': '',
                    'end': body.get('end','') or '',
                    'enabled': body.get('enabled') is not False,
                    # 是否占用预算：勾选后，本月未到期的这笔会先在预算进度条上以「待扣」占用额度
                    'budget': body.get('budget') is not False,
                    # 分期付款：0 = 一次性；>=2 = 总共分几期（金额存的是「每期金额」，等分）
                    'periods': _int(body.get('periods'), 0),
                    'createdAt': int(time.time()*1000),
                }
                items.append(it)
                save_recurring(items)
                self.json_resp(it,201)
            elif p == '/api/recurring/apply':
                added = maybe_apply_recurring()
                self.json_resp({'ok':True,'added':added})
            elif p == '/api/import':
                r = load_records(); incoming = body.get('records',[])
                if not incoming: self.json_resp({'error':'No records'},400); return
                current_id = next_id(r); now = int(time.time()*1000)
                existing = set(x.get('id') for x in r)
                for rec in incoming:
                    rid = rec.get('id')
                    if rid is not None and rid not in existing:
                        rec['id'] = rid
                    else:
                        while current_id in existing:
                            current_id += 1
                        rec['id'] = current_id; current_id += 1
                    if not rec.get('createdAt'):
                        rec['createdAt'] = now + rec['id']
                    # 撤回退款时恢复原支出的 refundedBy（删除时被清理）
                    if rec.get('type') == 'income' and rec.get('category') == REFUND_CAT and rec.get('refundOf') is not None:
                        for rr in r:
                            if rr.get('id') == rec.get('refundOf') and rr.get('type') == 'expense':
                                rr['refundedBy'] = rec['id']
                                break
                    r.append(rec)
                    existing.add(rec['id'])
                save_records(r); self.json_resp({'ok':True,'imported':len(incoming)})
            elif p == '/api/replace_all':
                # 整体替换（一键填充示例数据用）：先备份到回收站，再写入新记录
                incoming = body.get('records',[])
                if not incoming: self.json_resp({'error':'No records'},400); return
                r = load_records()
                if r:
                    _now_ms = int(time.time() * 1000)
                    for _x in r:
                        _x.setdefault('deletedAt', _now_ms)
                    save_trash(r)  # 同上：备份到撤回列表
                current_id = next_id([]); now = int(time.time()*1000)
                existing = set()
                for rec in incoming:
                    rid = rec.get('id')
                    if rid is not None and rid not in existing:
                        rec['id'] = rid
                    else:
                        while current_id in existing:
                            current_id += 1
                        rec['id'] = current_id; current_id += 1
                    if not rec.get('createdAt'):
                        rec['createdAt'] = now + rec['id']
                    existing.add(rec['id'])
                save_records(incoming)
                self.json_resp({'ok':True,'imported':len(incoming)})
            elif p == '/api/upgrade':
                self.upgrade_from_zip(body)
            elif p == '/api/import_xlsx':
                self.import_xlsx(body)
            elif p == '/api/undo':
                trash = load_trash()
                if not trash: self.json_resp({'error':'没有可撤回的记录'},400); return
                r = load_records()
                current_id = next_id(r)
                now = int(time.time()*1000)
                existing = set(x.get('id') for x in r)
                trashed = []
                for rec in trash:
                    rid = rec.get('id')
                    if rid is not None and rid not in existing:
                        rec['id'] = rid
                    else:
                        while current_id in existing:
                            current_id += 1
                        rec['id'] = current_id; current_id += 1
                    if not rec.get('createdAt'):
                        rec['createdAt'] = now + rec['id']
                    # 撤回退款时恢复原支出的 refundedBy（删除时被清理）
                    if rec.get('type') == 'income' and rec.get('category') == REFUND_CAT and rec.get('refundOf') is not None:
                        for rr in r:
                            if rr.get('id') == rec.get('refundOf') and rr.get('type') == 'expense':
                                rr['refundedBy'] = rec['id']
                                break
                    r.append(rec)
                    trashed.append(rec)
                    existing.add(rec['id'])
                save_records(r)
                save_trash([])
                self.json_resp({'ok':True,'restored':len(trashed)})
            elif p == '/api/clear':
                r = load_records()
                if r:
                    _now_ms = int(time.time() * 1000)
                    for _x in r:
                        _x.setdefault('deletedAt', _now_ms)
                    save_trash(r)  # 备份到回收站
                save_records([])
                self.json_resp({'ok':True})
            elif p == '/api/backup':
                name = do_backup()
                self.json_resp({'ok':True,'name':name})
            elif p == '/api/backup/import':
                _raw = b''
                try:
                    _b64 = body.get('file', '') or ''
                    if _b64.startswith('data:'):
                        _b64 = _b64.split(',', 1)[-1]
                    _raw = base64.b64decode(_b64) if _b64 else b''
                except Exception:
                    self.json_resp({'error': '文件内容损坏，无法解码'}, 400)
                    return
                ok, msg = import_backup_package(_raw)
                if ok:
                    self.json_resp({'ok': True, 'msg': msg})
                else:
                    self.json_resp({'error': msg}, 400)
            elif p == '/api/backup/restore':
                name = body.get('name','')
                if restore_backup(name):
                    self.json_resp({'ok':True})
                else:
                    self.json_resp({'error':'备份不存在'},404)
            else: self.json_resp({'error':'not found'},404)
        except Exception as e:
            try: self.send_response(500); self.send_cors(); self.send_header('Content-Type','text/plain'); self.end_headers(); self.wfile.write(str(e).encode('utf-8'))
            except Exception:
                pass

    def do_PUT(self):
        p = urllib.parse.urlparse(self.path).path; body = self.read_body()
        try:
            if p == '/api/records':
                r = load_records()
                rid = body.get('id')
                updated = False
                for rec in r:
                    if rec.get('id') == rid:
                        # 支持更新字段：note / category / amount / date（可选传入）
                        if 'note' in body:
                            rec['note'] = body.get('note','')
                        if 'category' in body:
                            rec['category'] = body.get('category')
                        if 'sub' in body:
                            rec['sub'] = body.get('sub','')
                        if 'amount' in body:
                            try:
                                amt = float(body.get('amount'))
                                if amt <= 0:
                                    self.json_resp({'error':'金额必须大于0'},400); return
                                rec['amount'] = round(amt, 2)
                            except (TypeError, ValueError):
                                self.json_resp({'error':'金额格式不对'},400); return
                        if 'date' in body:
                            new_date = body.get('date','')
                            # 简单校验 YYYY-MM-DD 格式
                            import re
                            if not re.match(r'^\d{4}-\d{2}-\d{2}$', new_date):
                                self.json_resp({'error':'日期格式不对'},400); return
                            rec['date'] = new_date
                        # 退款关联字段：refundOf = 原支出id（仅收入/退款类可设）；refundedBy = 反向指针
                        if 'refundOf' in body:
                            ro = body.get('refundOf')
                            if ro is None or ro == '':
                                # 解除关联：同时清理原支出上的 refundedBy
                                old_ro = rec.get('refundOf')
                                rec.pop('refundOf', None)
                                if old_ro is not None:
                                    for rr in r:
                                        if rr.get('id') == old_ro and rr.get('refundedBy') == rid:
                                            rr.pop('refundedBy', None)
                                            break
                            else:
                                try: ro = int(ro)
                                except (TypeError, ValueError):
                                    self.json_resp({'error':'关联ID格式不对'},400); return
                                # 目标必须是存在的支出记录
                                target = None
                                for rr in r:
                                    if rr.get('id') == ro and rr.get('type') == 'expense':
                                        target = rr; break
                                if target is None:
                                    self.json_resp({'error':'关联的支出不存在'},400); return
                                # 解除旧的 refundedBy 指针
                                old_ro = rec.get('refundOf')
                                if old_ro is not None and old_ro != ro:
                                    for rr in r:
                                        if rr.get('id') == old_ro and rr.get('refundedBy') == rid:
                                            rr.pop('refundedBy', None)
                                            break
                                rec['refundOf'] = ro
                                target['refundedBy'] = rid
                        if 'refundedBy' in body:
                            rb = body.get('refundedBy')
                            if rb is None or rb == '':
                                old_rb = rec.get('refundedBy')
                                rec.pop('refundedBy', None)
                                if old_rb is not None:
                                    for rr in r:
                                        if rr.get('id') == old_rb and rr.get('refundOf') == rid:
                                            rr.pop('refundOf', None)
                                            break
                            else:
                                try: rb = int(rb)
                                except (TypeError, ValueError):
                                    self.json_resp({'error':'关联ID格式不对'},400); return
                                target = None
                                for rr in r:
                                    if rr.get('id') == rb and rr.get('type') == 'income' and rr.get('category') == REFUND_CAT:
                                        target = rr; break
                                if target is None:
                                    self.json_resp({'error':'关联的退款不存在'},400); return
                                old_rb = rec.get('refundedBy')
                                if old_rb is not None and old_rb != rb:
                                    for rr in r:
                                        if rr.get('id') == old_rb and rr.get('refundOf') == rid:
                                            rr.pop('refundOf', None)
                                            break
                                rec['refundedBy'] = rb
                                target['refundOf'] = rid
                        updated = True
                        break
                if not updated:
                    self.json_resp({'error':'Record not found'},404); return
                save_records(r)
                self.json_resp({'ok':True,'note':body.get('note',rec.get('note','')),'category':body.get('category',rec.get('category','')),'amount':rec.get('amount')})
            elif p == '/api/trash':
                # 回收站整体重写（恢复单条/批量恢复/彻底删除/清空都走这里）
                # body.records 必须是数组；只保留必要字段，防止被塞进奇怪的东西
                recs = body.get('records')
                if recs is None or not isinstance(recs, list):
                    self.json_resp({'error':'records 必须是数组'},400); return
                clean = []
                for x in recs:
                    if not isinstance(x, dict): continue
                    try:
                        clean.append({
                            'id': x.get('id'),
                            'date': x.get('date',''),
                            'type': x.get('type','expense'),
                            'category': x.get('category',''),
                            'sub': x.get('sub',''),
                            'amount': float(x.get('amount') or 0),
                            'note': x.get('note',''),
                            'createdAt': x.get('createdAt') or int(time.time()*1000),
                            # 删除时间：自动清理靠它算保留期，写回时不能丢
                            'deletedAt': int(x.get('deletedAt') or 0) or int(time.time()*1000),
                        })
                    except (TypeError, ValueError):
                        continue
                save_trash(clean)
                self.json_resp({'ok':True,'count':len(clean)})
            elif p == '/api/recurring':
                items = load_recurring()
                rid = body.get('id')
                for it in items:
                    if it.get('id') == rid:
                        for k in ('name','category','sub','note','freq','start','end','month','day','enabled','lastAuto','budget'):
                            if k in body:
                                it[k] = body.get(k)
                        if 'periods' in body:
                            try:
                                it['periods'] = int(body.get('periods') or 0)
                            except (TypeError, ValueError):
                                it['periods'] = 0
                        if 'type' in body:
                            it['type'] = 'income' if body.get('type') == 'income' else 'expense'
                        if 'amount' in body:
                            try:
                                amt = float(body.get('amount'))
                            except (TypeError, ValueError):
                                self.json_resp({'error':'金额格式不对'},400); return
                            if amt <= 0:
                                self.json_resp({'error':'金额必须大于0'},400); return
                            it['amount'] = round(amt, 2)
                        save_recurring(items)
                        self.json_resp(it)
                        return
                self.json_resp({'error':'固定账单不存在'},404)
            else:
                self.json_resp({'error':'not found'},404)
        except Exception as e:
            try: self.send_response(500); self.send_cors(); self.send_header('Content-Type','text/plain'); self.end_headers(); self.wfile.write(str(e).encode('utf-8'))
            except Exception:
                pass

    def do_DELETE(self):
        p = urllib.parse.urlparse(self.path).path; body = self.read_body()
        try:
            if p == '/api/records':
                ids = body.get('ids',[])
                if not ids: self.json_resp({'error':'No ids'},400); return
                r = load_records(); id_set = set(ids)
                deleted_records = [x for x in r if x.get('id') in id_set]
                new_r = [x for x in r if x.get('id') not in id_set]
                # 清理反向引用：删除退款时清掉原支出的 refundedBy；删除支出时清掉退款的 refundOf
                for x in new_r:
                    if x.get('refundOf') in id_set:
                        x['refundOf'] = None
                    if x.get('refundedBy') in id_set:
                        x['refundedBy'] = None
                # 保存到回收站（打上删除时间戳，供自动清理判断）
                if deleted_records:
                    now_ms = int(time.time() * 1000)
                    for x in deleted_records:
                        if not x.get('deletedAt'):
                            x['deletedAt'] = now_ms
                    trash = load_trash()
                    trash.extend(deleted_records)
                    save_trash(trash)
                save_records(new_r)
                self.json_resp({'ok':True,'deleted':len(deleted_records),'deletedRecords':deleted_records})
            elif p == '/api/recurring':
                rid = body.get('id')
                items = load_recurring()
                keep = [x for x in items if x.get('id') != rid]
                if len(keep) == len(items):
                    self.json_resp({'error':'固定账单不存在'},404); return
                save_recurring(keep)
                self.json_resp({'ok':True,'deleted':1})
            elif p == '/api/backup':
                name = body.get('name','')
                if delete_backup(name):
                    self.json_resp({'ok':True})
                else:
                    self.json_resp({'error':'备份不存在'},404)
            else: self.json_resp({'error':'not found'},404)
        except Exception as e:
            try: self.send_response(500); self.send_cors(); self.send_header('Content-Type','text/plain'); self.end_headers(); self.wfile.write(str(e).encode('utf-8'))
            except Exception:
                pass

def backup_ledger_before_upgrade(bak_dir):
    """升级前自动导出一份「数据备份包」（2026-10-18 主人要求）。

    主人：「升级那里也改一下，把他改成升级时自动导出一份数据备份包就行了」

    直接用 export_backup_package()，导出来的就是设置页「导出备份包」同一个格式
    （小爱记账备份_<时间戳>.zip + 备份说明.txt），
    以后数据出问题可以直接「设置 → 数据备份 → 从文件恢复」选它。

    放在程序备份同一个目录（data_backup\\upgrade_<stamp>\\），一处就能找全。
    失败不影响升级主流程，只在结果里标一下。
    """
    try:
        os.makedirs(bak_dir, exist_ok=True)
        return export_backup_package(out_dir=bak_dir)
    except Exception:
        return ''


def _program_blob_problem(name, blob):
    """检查即将写入的程序文件内容是不是明显不对。

    返回 None 表示没问题；否则返回一句中文原因。
    只做最低限度的形式校验，不会误杀正常文件。
    """
    low = (name or '').lower()
    try:
        if not blob:
            return '文件是空的'
        if low.endswith('.html') or low.endswith('.htm'):
            head = blob[:4096].lower()
            if b'<html' not in head and b'<!doctype' not in head:
                return '不是一个网页文件'
        elif low.endswith('.py'):
            if b'def ' not in blob and b'import ' not in blob:
                return '不是一个 Python 程序'
        elif low.endswith('.js'):
            if b'function' not in blob and b'=>' not in blob:
                return '不是一个 JS 文件'
        elif low.endswith('.exe'):
            if not blob.startswith(b'MZ'):
                return '不是 Windows 程序'
        elif low.endswith('.ico'):
            if not blob.startswith(b'\x00\x00\x01\x00'):
                return '不是图标文件'
    except Exception:
        pass
    return None


def _write_staged_exe(dest, blob):
    """把新 exe 写成 `<原 exe>.new`，供重启时换上。

    返回 (ok, reason)。

    ⚠️ 2026-10-18 主人实际遇到的 bug：
    `[Errno 13] Permission denied: '...小爱记账.exe.new'`

    现场证据：用户的目录里连续升级了三次（02:19:59 / 02:20:19 / 02:20:36），
    `.new` 的 mtime 停在**第一次**，后两次都 Permission denied。
    而事后去检测，`.new` 又完全不锁、能重命名——说明是**瞬时占用**，
    典型来源：① 杀毒软件/Windows Defender 刚扫完旧 `.new` 还没放开；
    ② 上一轮已经 spawn 的助手正在 `move` 它；③ 资源管理器正在预览它。

    旧代码 `open(dest+'.new','wb')` 一次失败就直接放弃，对用户来说就是
    “升级失败”——但重新点一次又好了，体验很差。

    现在改成：
      1. **先清掉可能残留的旧 `.new`**（里面是上一个版本，本来就要被替换）
      2. **重试写入**，每次 0.6 秒，最多 12 次（约 7 秒）
      3. 全部失败时，**换一个文件名**（`.exe.new2`）再试一次——
         助手本来就是用 `*.exe.new` 通配符找的，但换个名字能绕开
         “某个名字被锁死”的轻微情况；实在不行才报错给用户
    """
    import time as _t

    targets = [dest + '.new', dest + '.new2']
    last_err = None

    for tf in targets:
        # ① 先尝试清掉同名残留（占着就下一轮再说）
        for _ in range(3):
            if not os.path.isfile(tf):
                break
            try:
                os.remove(tf)
                break
            except Exception:
                _t.sleep(0.4)

        # ② 重试写入
        for i in range(8):
            try:
                with open(tf, 'wb') as f:
                    f.write(blob)
                # 写完校验一下，防止写一半
                if os.path.getsize(tf) == len(blob):
                    return True, None
                last_err = '写入不完整'
            except PermissionError as e:
                last_err = e
            except Exception as e:
                last_err = e
                break  # 不是占用问题，换名字也没用
            _t.sleep(0.5)

    # ③ 都不行：把话说清楚，告诉用户怎么办
    msg = str(last_err)
    if isinstance(last_err, PermissionError) or 'Permission denied' in msg or 'Error 13' in msg:
        msg = ('文件被占用（杀毒软件或上一轮重启助手正在处理）。'
               '请稍等几秒再点一次升级，或者先关掉程序再升级')
    return False, msg


def _upgrade_helper_bat():
    """升级助手：等程序退出 -> 换上 .new -> **并且只在真的换上了才启动**。

    ⚠⚠ 2026-10-18 第六版，改了一个**非常严重的错误**：

    主人反馈：「只要点击升级选完文件后，如果不点击“立即重启”就一直弹窗，一直跳」
    —— 至少二三十个浏览器窗口。

    根因是前几版的设计错得离谱：
      助手被拉起来后**根本不等主人点重启**，只要发现 `.new` 存在、
      而旧进程刚好还没跑（或者已经被杀），它就 `explorer.exe` 启动程序；
      而**程序每次启动都会自动开一个浏览器窗口**。
      于是：重试 + 多开助手 + 程序自启动浏览器 = 无限弹窗。

    第六版的铁律：
      1. **助手只在真的把 `.new` 换到 `.exe` 成功后**才启动程序；
         换不上就默默退出，绝不启动、绝不刷窗口。
      2. **`launchDone` 标记**：本目录下写过 `_helper.launched`
         就再也不再启动（避免反复拉起）。
      3. **超时更短**：只等 120 秒，等不到就放弃（不启动）。
      4. 副本文件（`.new` / `.new2`）在启动前先删干净。
    """
    w = [
        '@echo off',
        'setlocal',
        'cd /d "%~dp0"',
        'set "LOG=%~dp0_upgrade_helper.log"',
        'set "LOCK=%~dp0_helper.lock"',
        'set "DONE=%~dp0_helper.launched"',
        'echo [%date% %time%] start >> "%LOG%"',
        '',
        'rem ---- single instance: wait 2s, then decide ----',
        'if exist "%LOCK%" (',
        '  ping -n 3 127.0.0.1 >nul',
        '  if exist "%LOCK%" (',
        '    echo [%date% %time%] another helper busy, exit >> "%LOG%"',
        '    exit /b',
        '  )',
        ')',
        'echo running > "%LOCK%"',
        '',
        'rem ---- locate staged exe (no hardcoded name) ----',
        'set "ORIG="',
        'set "NEWP="',
        'for %%F in ("*.exe.new") do if not defined ORIG set "ORIG=%%~dpnF"',
        'for %%F in ("*.exe.new") do if not defined NEWP set "NEWP=%%~fF"',
        'if not defined ORIG for %%F in ("*.exe.new2") do set "ORIG=%%~dpnF"',
        'if not defined ORIG for %%F in ("*.exe.new2") do set "NEWP=%%~fF"',
        'if not defined ORIG (',
        '  echo [%date% %time%] nothing staged, done >> "%LOG%"',
        '  goto cleanup',
        ')',
        'echo [%date% %time%] target=%ORIG% >> "%LOG%"',
        '',
        'rem ---- already launched once? never launch again ----',
        'if exist "%DONE%" (',
        '  echo [%date% %time%] already launched before, only swap >> "%LOG%"',
        ')',
        '',
        'rem ---- wait for the running exe to release its file, then swap ----',
        'set /a TRIES=0',
        ':retry',
        'if not exist "%NEWP%" goto swapped',
        'set /a TRIES+=1',
        'if %TRIES% gtr 120 goto giveup',
        'move /y "%NEWP%" "%ORIG%" >nul 2>&1',
        'if exist "%NEWP%" (',
        '  ping -n 2 127.0.0.1 >nul',
        '  goto retry',
        ')',
        '',
        ':swapped',
        'echo [%date% %time%] swapped ok >> "%LOG%"',
        'if exist "%DONE%" (',
        '  echo [%date% %time%] launch suppressed >> "%LOG%"',
        '  goto cleanup',
        ')',
        'echo done > "%DONE%"',
        'echo [%date% %time%] launching >> "%LOG%"',
        'explorer.exe "%ORIG%"',
        'ping -n 4 127.0.0.1 >nul',
        'goto cleanup',
        '',
        ':giveup',
        'echo [%date% %time%] giveup, file still locked, NOT launching >> "%LOG%"',
        'rem keep the staged copy so the user can retry later; just release the lock',
        'del /f /q "%LOCK%" >nul 2>&1',
        'echo [%date% %time%] giveup done >> "%LOG%"',
        'del /f /q "%~f0" >nul 2>&1',
        'endlocal',
        'exit /b',
        '',
        ':cleanup',
        'del /f /q "*.exe.new" >nul 2>&1',
        'del /f /q "*.exe.new2" >nul 2>&1',
        'del /f /q "%LOCK%" >nul 2>&1',
        'echo [%date% %time%] done >> "%LOG%"',
        'del /f /q "%~f0" >nul 2>&1',
        'endlocal',
        'exit /b',
    ]
    return '\r\n'.join(w) + '\r\n'


def _spawn_upgrade_helper(helper_path):
    """把升级助手在**后台脱离启动**，不阻塞当前进程。

    助手会一直等旧 exe 退出，然后把 .new 换上、重新启动。
    用 DETACHED_PROCESS 让它不跟当前进程共生死（否则进程一退它也没了）。
    成功返回 True。
    """
    import subprocess
    if not helper_path or not os.path.isfile(helper_path):
        return False
    try:
        flags = 0
        # ⚠⚠ 2026-10-18 第七版修正：**不要用 DETACHED_PROCESS**。
        #   DETACHED = “完全不要控制台”，这会让子进程里的控制台程序
        #   （`timeout` 之类）自己新开一个**可见窗口** ——
        #   主人看到的两个黑窗就是这么来的。
        #   改用 CREATE_NO_WINDOW：同样不显示窗口，但是给了一个隐藏控制台，
        #   子进程不会再去新建可见窗口。
        if hasattr(subprocess, 'CREATE_NO_WINDOW'):
            flags |= subprocess.CREATE_NO_WINDOW
        else:
            flags |= subprocess.DETACHED_PROCESS
        if hasattr(subprocess, 'CREATE_NEW_PROCESS_GROUP'):
            flags |= subprocess.CREATE_NEW_PROCESS_GROUP
        # 双保险：再显式要求隐藏窗口
        si = None
        try:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            si.wShowWindow = subprocess.SW_HIDE
        except Exception:
            si = None
        subprocess.Popen(
            ['cmd.exe', '/c', helper_path],
            cwd=EXE_DIR,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=flags,
            startupinfo=si,
            close_fds=True,
        )
        return True
    except Exception:
        return False


def _project_launcher():
    """分享包里的「启动.bat」：双击即用"""
    w = [
        '@echo off',
        'chcp 65001 >nul',
        'cd /d "%~dp0"',
        'if exist "小爱记账.exe" (',
        '  start "" "小爱记账.exe"',
        ') else (',
        '  echo 找不到小爱记账.exe，请确认解压完整喵～',
        '  pause',
        ')',
    ]
    return '\r\n'.join(w) + '\r\n'


def _project_readme():
    """分享包里的「使用说明.txt」"""
    L = [
        '小爱记账 · 空白账本喵',
        '========================',
        '',
        '【怎么用】',
        '1. 双击「启动.bat」（或直接双击「小爱记账.exe」）',
        '2. 浏览器会自动打开 http://127.0.0.1:8899/',
        '3. 开始记账喵～',
        '',
        '【数据在哪】',
        '都在本文件夹的 data\\ 目录里：',
        '  records.json   记账记录',
        '  trash.json     回收站',
        '  budget.json    预算设置',
        '  recurring.json 固定账单',
        '  settings.json  分类配置',
        '想备份直接复制整个 data\\ 文件夹就行喵。',
        '',
        '【要注意】',
        '- 只在本机运行，不联网，数据不上传到任何地方',
        '- 默认端口 8899（想换端口就在程序目录放 port.txt），打不开先看看是不是被别的程序占了',
        '- 换电脑 / 换U盘：整个文件夹拷过去就能用，数据跟着走',
        '',
        '【它是个空的】',
        '这份是干净的空白账本，里面没有任何别人的账单喵。',
        '第一次打开就是全新的开始～',
    ]
    return '\r\n'.join(L) + '\r\n'


def _clean_upgrade_debris():
    """程序启动时清理上一次升级留下的残渣。

    为什么要做：
      ① 主人可能没有点「立即重启」，而是自己手动关了再开的程序 ->
         `.new` 就永远留在那里，下次升级又去写它、又撞上占用。
      ② 助手如果被任务管理器杀掉，`_helper.lock` 会残留 ->
         以后每次启动助手都 `if exist "%LOCK%" exit /b` 直接退出，
         自动重启就彻底失效了。

    做法：启动时如果发现 `.new` / `.new2` 已经存在，
    先比较是不是同一个 exe（大小 + 首尾字节），是就删掉；不是就改名存成 .old 免得丢。
    最后把 `_helper.lock` 和过期的 bat 一并清掉。
    """
    try:
        cur = sys.executable if getattr(sys, 'frozen', False) else ''
        cur_sz = os.path.getsize(cur) if cur and os.path.isfile(cur) else 0

        for suf in ('.new', '.new2'):
            for n in os.listdir(EXE_DIR):
                if not n.lower().endswith('.exe' + suf):
                    continue
                p = os.path.join(EXE_DIR, n)
                try:
                    if cur_sz and os.path.getsize(p) == cur_sz:
                        os.remove(p)          # 跟自己一样大，大概率已经是换上去后的同一份
                    else:
                        # 不确定：留着但改名，不要跟下次升级撞名字
                        try:
                            os.remove(p + '.old')
                        except Exception:
                            pass
                        os.rename(p, p + '.old')
                except Exception:
                    pass

        # 残留锁：程序都已经重新启动了，旧锁肯定是死的
        for n in ('_helper.lock', '_helper.launched'):
            p = os.path.join(EXE_DIR, n)
            if os.path.isfile(p):
                try:
                    os.remove(p)
                except Exception:
                    pass
    except Exception:
        pass


# ── 端口配置（2026-10-28）────────────────────────────────────────────
#   同一台电脑可以放多份账本（比如一份正式账、一份空白账本拿来试功能）。
#   每份账本用不同端口，浏览器地址就不同，互不踩脚。
#   解析顺序：环境变量 XIAOAI_LEDGER_PORT > 程序目录 port.txt > 默认 8899
DEFAULT_PORT = 8899


def _read_port_file():
    try:
        p = os.path.join(EXE_DIR, 'port.txt')
        if os.path.isfile(p):
            with open(p, 'r', encoding='utf-8-sig', errors='ignore') as f:
                for line in f:
                    s = line.strip()
                    if not s or s.startswith('#'):
                        continue
                    return int(s)
    except Exception:
        pass
    return None


def _resolve_port():
    for raw_v in (os.environ.get('XIAOAI_LEDGER_PORT'), _read_port_file()):
        try:
            if raw_v not in (None, ''):
                n = int(str(raw_v).strip())
                if 1024 <= n <= 65535:
                    return n
        except Exception:
            pass
    return DEFAULT_PORT


def _self_exe_path():
    try:
        if getattr(sys, 'frozen', False):
            return os.path.abspath(sys.executable)
        return os.path.abspath(__file__)
    except Exception:
        return ''


def _same_path(a, b):
    try:
        if not a or not b:
            return False
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except Exception:
        return False


def _pid_exe_path(pid):
    try:
        import subprocess
        out = subprocess.check_output(
            ['powershell', '-NoProfile', '-Command',
             '(Get-Process -Id %d -ErrorAction SilentlyContinue).Path' % int(pid)],
            stderr=subprocess.DEVNULL)
        return out.decode('gbk', 'ignore').strip()
    except Exception:
        return ''


def _pid_cmdline(pid):
    """取进程命令行（脚本模式下用它判断是不是同一份账本）。"""
    try:
        import subprocess
        out = subprocess.check_output(
            ['powershell', '-NoProfile', '-Command',
             '(Get-CimInstance Win32_Process -Filter "ProcessId=%d" | '
             'Select-Object -ExpandProperty CommandLine)' % int(pid)],
            stderr=subprocess.DEVNULL)
        return out.decode('gbk', 'ignore').strip()
    except Exception:
        return ''


def _is_same_instance(pid):
    """这个 pid 是不是「同一个文件夹里的自己」。
    exe 模式比对可执行文件路径；脚本模式再比对命令行里的脚本路径。
    """
    me = _self_exe_path()
    if not me:
        return False
    if _same_path(_pid_exe_path(pid), me):
        return True
    if not getattr(sys, 'frozen', False):
        cmd = _pid_cmdline(pid)
        if cmd and os.path.normcase(me) in os.path.normcase(cmd):
            return True
    return False


def _msgbox(text, title='小爱记账'):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, text, title, 0x40)
    except Exception:
        pass


def main():
    PORT = _resolve_port()
    # ⚠ 2026-10-18：先记下「这次是不是升级后的自动重启」，再去做清残。
    #   因为 _clean_upgrade_debris() 会把 _helper.launched 删掉，
    #   如果先清再判，就永远判不到 → 浏览器又开了 → 又弹窗。
    try:
        _post_upgrade_launch = os.path.isfile(os.path.join(EXE_DIR, '_helper.launched'))
    except Exception:
        _post_upgrade_launch = False

    # 只清理「同一个文件夹」里占着本端口的旧进程（同一份账本重开时接管）
    #   ⚠ 2026-10-18：以前端口写死 8899，测试实例会把正式程序杀掉。
    #   ⚠ 2026-10-28：同机可以有多份账本（不同文件夹）。路径不同就是不同账本，
    #   绝不能互相杀 —— 只有 exe 路径和本次完全一致才清理。
    try:
        import subprocess
        out = subprocess.check_output('netstat -ano | findstr :%d' % PORT,
                                      shell=True, text=True)
        for line in out.strip().split('\n'):
            parts = line.strip().split()
            if len(parts) >= 5 and 'LISTENING' in line:
                pid = int(parts[-1])
                if pid and pid != os.getpid():
                    if _is_same_instance(pid):
                        subprocess.call('taskkill /f /pid %d' % pid, shell=True,
                                        stdout=subprocess.DEVNULL,
                                        stderr=subprocess.DEVNULL)
    except Exception:
        pass
    time.sleep(0.5)

    # 启动时自动备份（每天最多一次）
    try:
        today = time.strftime('%Y%m%d')
        auto = False
        if os.path.isdir(BACKUP_DIR):
            for n in os.listdir(BACKUP_DIR):
                if n.startswith(today):
                    auto = True
                    break
        if not auto:
            do_backup()
    except: pass

    # 启动时顺手清一次回收站（按保留天数），不用等用户打开回收站
    try:
        _n = auto_purge_trash()
        if _n:
            print('[trash] 自动清理了 %d 条过期记录' % _n)
    except Exception as _e:
        print('[trash] 自动清理失败：', _e)

    # 启动服务器前，先补齐到期的固定账单（补最近 3 个月）
    try:
        maybe_apply_recurring()
    except Exception:
        pass

    # ⚠ 2026-10-18：启动时清理升级残留（否则会卡住以后的自动重启）
    try:
        _clean_upgrade_debris()
    except Exception:
        pass

    # 启动服务器
    #   2026-10-28：端口被另一份账本占着时别静默崩掉 —— 弹提示说明怎么改端口
    try:
        srv = ThreadingHTTPServer(('127.0.0.1', PORT), Handler)
    except OSError as _e:
        _msg = ('端口 %d 起不来（%s）。\n\n'
                '如果这是另一份账本（不在同一个文件夹），请在本程序目录新建 '
                'port.txt，写一个没用过的端口号（比如 8907），再重新打开。\n\n'
                '如果是同一份账本重复打开了，关掉多余的那个再试。' % (PORT, _e))
        print(_msg)
        _msgbox(_msg)
        return
    srv.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)

    # 子线程开浏览器，不阻塞主线程
    # ⚠⚠ 2026-10-18：**只有主人手动双击启动时才开浏览器**。
    #   以前无条件开 → 升级/重启循环时每启动一次就弹一个窗口，
    #   实测弹了二三十次（主人的原话：「一直弹窗，一直跳」）。
    #   现在：升级后自动重启的那一次不开（主人本来就有页面）。
    def _open_browser():
        time.sleep(0.5)
        try: webbrowser.open('http://127.0.0.1:%d/' % PORT)
        except: pass
    if not _post_upgrade_launch and not os.environ.get('XIAOAI_NO_BROWSER'):
        threading.Thread(target=_open_browser, daemon=True).start()
    else:
        print('[browser] 升级后自动重启，不打开浏览器' if _post_upgrade_launch else '[browser] 已禁用')

    # 主线程serve_forever，永不退出
    srv.serve_forever()

if __name__ == '__main__':
    main()
