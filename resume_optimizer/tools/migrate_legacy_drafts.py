# -*- coding: utf-8 -*-
"""存量草稿降级工具：把「真实模板 + 富文本」草稿降级为纯经典草稿。

背景（产品决策）：编辑器已彻底移除真实 Word 模板，只保留 8 套内置模板；
真实模板仅用于首页展示与下载 Word 原件。库里仍可能残留绑定真实模板、
并带 richtext_mode / richtext_edits 的旧草稿，新版前端会忽略这些字段，
本脚本提供一个可审计、可回滚的降级手段（不动草稿本身，只精简 data_json）。

安全约定：
- 默认 --dry-run：只读扫描并打印报告，绝不写库。
- --backup：把受影响草稿的原始 data_json 导出到备份目录。
- --apply 必须同时给 --yes 才真正写库；写库前自动备份原始 JSON。
- --rollback <备份目录> 从备份还原（同样需要 --yes）。

用法：
  python tools/migrate_legacy_drafts.py                    # 只读诊断（默认行为）
  python tools/migrate_legacy_drafts.py --backup           # 导出备份
  python tools/migrate_legacy_drafts.py --apply --yes      # 真正降级
  python tools/migrate_legacy_drafts.py --rollback D:/tmp/legacy_drafts_backup --yes
"""
import argparse
import datetime
import hashlib
import json
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

# 数据库访问走项目既有 utils/db.py（.env 的 DB_* 配置），不硬编码连接参数
from utils.db import DB_HOST, DB_NAME, DB_PORT, get_conn  # noqa: E402

DEFAULT_BACKUP_DIR = "D:/tmp/legacy_drafts_backup"

# 降级时要从 data_json 里摘掉的键：
# richtext_mode / richtext_edits 是旧富文本模式的开关与就地编辑快照；
# template_id / real_template_id 是旧版本写进 JSON 的真实模板绑定（见 db._extract_template_id_from_data）。
DROP_KEYS = ("richtext_mode", "richtext_edits", "template_id", "real_template_id")

# 经典模式（editor-v3.js normalizeData）必需的顶层结构
CLASSIC_TOP_KEYS = ("resume", "theme")


def _setup_stdout():
    """中文报告在 Windows 控制台（cp936）下也能原样输出。"""
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001 - 老解释器或重定向场景，忽略即可
        pass


def _loads(raw):
    """解析 data_json，失败返回 (None, 错误信息)。"""
    try:
        data = json.loads(raw) if raw else {}
    except (ValueError, TypeError) as exc:
        return None, "data_json 不是合法 JSON：%s" % exc
    if not isinstance(data, dict):
        return None, "data_json 顶层不是对象：%s" % type(data).__name__
    return data, ""


def _classic_ready(data):
    """判断该草稿能否在经典模式下正常打开，返回 (是否可打开, 说明)。

    兼容 editor-v3.js 的两套模型：
    - v3：{v:3, resume:{...}, theme:{...}}；
    - 旧模型：{basic / modules ...}，前端会尽力迁移。
    """
    missing = [k for k in CLASSIC_TOP_KEYS if not isinstance(data.get(k), (dict,))]
    if not missing:
        resume = data.get("resume") or {}
        if not isinstance(resume.get("basic"), dict) and not isinstance(resume.get("modules"), dict):
            return False, "resume 里既没有 basic 也没有 modules（经典模式无可渲染内容）"
        theme_ok = bool(data.get("theme"))
        return True, "v=%s，含 resume/theme%s" % (
            data.get("v"), "" if theme_ok else "（theme 为空，将回落到默认皮肤）")
    if data.get("basic") or data.get("modules"):
        return True, "旧模型草稿（basic/modules），前端会自动迁移到 v3"
    return False, "缺少经典模式必需字段：%s（data_json 顶层键：%s）" % (
        "/".join(missing), ",".join(sorted(data.keys())))


def _pick_keys(data):
    """待移除的键（只取实际存在的，避免空操作）。"""
    return [k for k in DROP_KEYS if k in data]


def _sha256(text):
    return hashlib.sha256((text or "").encode("utf-8")).hexdigest()


def _now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def fetch_drafts(draft_ids=None):
    """只读拉取草稿行（SELECT，不写库）。"""
    sql = ("SELECT id, user_id, template_id, title, data_json, created_at, updated_at "
           "FROM resume_drafts")
    args = []
    if draft_ids:
        sql += " WHERE id IN (%s)" % ",".join(["%s"] * len(draft_ids))
        args = list(draft_ids)
    sql += " ORDER BY id ASC"
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, args)
            return cur.fetchall()


def analyze(rows):
    """把原始行整理成诊断条目列表。"""
    items = []
    for r in rows:
        data, err = _loads(r.get("data_json"))
        item = {
            "row": r,
            "ok": data is not None,
            "error": err,
            "data": data or {},
            "drop": _pick_keys(data) if data else [],
        }
        if data is not None:
            item["ready"], item["ready_note"] = _classic_ready(data)
            edits = data.get("richtext_edits")
            item["rt_edits_n"] = len(edits) if isinstance(edits, dict) else 0
        else:
            item["ready"], item["ready_note"] = False, "解析失败，无法判断"
            item["rt_edits_n"] = 0
        # 受影响 = data_json 里还留着旧字段，或表列 template_id 仍指向真实模板
        item["affected"] = bool(item["drop"]) or r.get("template_id") is not None
        items.append(item)
    return items


def print_line(text=""):
    print(text)


def report(items, mode):
    """打印逐条明细 + 汇总。"""
    title = {
        "dry-run": "DRY-RUN 只读诊断",
        "backup": "备份导出",
        "apply": "降级执行",
        "rollback": "回滚还原",
    }.get(mode, mode)
    print_line("=" * 74)
    print_line("存量草稿降级报告 · %s · %s" % (title, _now()))
    print_line("数据库：%s @ %s:%s（连接参数来自 .env 的 DB_* 配置）" % (DB_NAME, DB_HOST, DB_PORT))
    print_line("=" * 74)
    for it in items:
        r = it["row"]
        print_line("")
        print_line("[草稿 %s]" % r["id"])
        print_line("  user_id        : %s" % r.get("user_id"))
        print_line("  template_id    : %s" % ("—" if r.get("template_id") is None else r["template_id"]))
        print_line("  title          : %s" % (r.get("title") or ""))
        print_line("  richtext_mode  : %s" % (it["data"].get("richtext_mode")
                                              if it["ok"] else "（解析失败）"))
        print_line("  richtext_edits : %s 个键" % it["rt_edits_n"])
        print_line("  创建时间       : %s" % r.get("created_at"))
        print_line("  updated_at     : %s" % r.get("updated_at"))
        if it["ok"]:
            print_line("  data_json 顶层键: %s" % ",".join(it["data"].keys()))
            print_line("  经典模式可打开 : %s（%s）" % ("是" if it["ready"] else "否", it["ready_note"]))
            print_line("  待移除字段     : %s" % (",".join(it["drop"]) if it["drop"] else "无"))
        else:
            print_line("  经典模式可打开 : 否（%s）" % it["error"])
        print_line("  受影响         : %s" % ("是" if it["affected"] else "否"))


def summary(items):
    """汇总统计。"""
    total = len(items)
    affected = [it for it in items if it["affected"]]
    ready = [it for it in affected if it["ready"]]
    blocked = [it for it in affected if not it["ready"]]
    bad = [it for it in items if not it["ok"]]
    key_counter = {}
    for it in affected:
        for k in it["drop"]:
            key_counter[k] = key_counter.get(k, 0) + 1
    print_line("")
    print_line("-" * 74)
    print_line("汇总：草稿总数 %d，受影响 %d，不受影响 %d" % (total, len(affected), total - len(affected)))
    print_line("      降级后可在经典模式打开 %d，仍打不开 %d（需人工处理）" % (len(ready), len(blocked)))
    print_line("      受影响草稿 id：%s" % (",".join(str(it["row"]["id"]) for it in affected) or "无"))
    if key_counter:
        print_line("      待移除字段分布：%s" % "，".join(
            "%s×%d" % (k, key_counter[k]) for k in DROP_KEYS if k in key_counter))
    if bad:
        print_line("      !! data_json 解析失败：%s" % ",".join(str(it["row"]["id"]) for it in bad))
    print_line("-" * 74)
    return affected


def write_backup(items, backup_dir):
    """把受影响草稿的原始 data_json 原样导出（含清单与 sha256，便于审计）。"""
    os.makedirs(backup_dir, exist_ok=True)
    manifest = {
        "tool": "tools/migrate_legacy_drafts.py",
        "created_at": _now(),
        "database": "%s@%s:%s" % (DB_NAME, DB_HOST, DB_PORT),
        "note": "data_json 为原始内容原样保存，可用 --rollback 还原",
        "drafts": [],
    }
    paths = []
    for it in items:
        r = it["row"]
        raw = r.get("data_json") or ""
        fname = "draft_%s.json" % r["id"]
        path = os.path.join(backup_dir, fname)
        payload = {
            "draft_id": r["id"],
            "user_id": r.get("user_id"),
            "template_id": r.get("template_id"),
            "title": r.get("title"),
            "created_at": str(r.get("created_at") or ""),
            "updated_at": str(r.get("updated_at") or ""),
            "sha256": _sha256(raw),
            "data_json": raw,
        }
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)
        paths.append(path)
        manifest["drafts"].append({
            "draft_id": r["id"],
            "user_id": r.get("user_id"),
            "template_id": r.get("template_id"),
            "file": fname,
            "sha256": payload["sha256"],
            "dropped_keys": it["drop"],
        })
    manifest_path = os.path.join(backup_dir, "manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
    return manifest_path, paths


def apply_downgrade(items):
    """在同一事务里逐条移除旧字段；任一条失败则整体回滚。"""
    changed = []
    with get_conn() as conn:  # get_conn 已 autocommit=False + commit/rollback
        with conn.cursor() as cur:
            for it in items:
                r = it["row"]
                if not it["drop"]:
                    print_line("  [草稿 %s] data_json 无待移除字段，跳过" % r["id"])
                    continue
                before = it["data"]
                after = {k: v for k, v in before.items() if k not in it["drop"]}
                after_json = json.dumps(after, ensure_ascii=False, separators=(",", ":"))
                cur.execute("UPDATE resume_drafts SET data_json=%s WHERE id=%s",
                            (after_json, r["id"]))
                changed.append({
                    "draft_id": r["id"],
                    "before_keys": list(before.keys()),
                    "after_keys": list(after.keys()),
                    "removed": it["drop"],
                    "before_bytes": len(r.get("data_json") or ""),
                    "after_bytes": len(after_json),
                })
                print_line("")
                print_line("  [草稿 %s] 变更前 -> 变更后" % r["id"])
                print_line("    顶层键 : %s  ->  %s" % (",".join(before.keys()), ",".join(after.keys())))
                print_line("    移除   : %s" % ",".join(it["drop"]))
                print_line("    体积   : %d 字节  ->  %d 字节" % (
                    len(r.get("data_json") or ""), len(after_json)))
    return changed


def rollback(backup_dir):
    """按备份目录还原 data_json。"""
    manifest_path = os.path.join(backup_dir, "manifest.json")
    if not os.path.isfile(manifest_path):
        print_line("找不到备份清单：%s" % manifest_path)
        return False
    with open(manifest_path, encoding="utf-8") as fh:
        manifest = json.load(fh)
    entries = manifest.get("drafts") or []
    if not entries:
        print_line("备份清单为空（%s）" % manifest_path)
        return False
    print_line("备份创建时间：%s" % manifest.get("created_at"))
    print_line("备份来源数据库：%s" % manifest.get("database"))
    restored = 0
    with get_conn() as conn:
        with conn.cursor() as cur:
            for e in entries:
                path = os.path.join(backup_dir, e.get("file") or ("draft_%s.json" % e["draft_id"]))
                if not os.path.isfile(path):
                    print_line("  [草稿 %s] 备份文件缺失：%s，跳过" % (e["draft_id"], path))
                    continue
                with open(path, encoding="utf-8") as fh:
                    payload = json.load(fh)
                raw = payload.get("data_json") or ""
                if payload.get("sha256") and _sha256(raw) != payload["sha256"]:
                    print_line("  [草稿 %s] sha256 校验不一致，拒绝还原（文件可能被改动）" % e["draft_id"])
                    continue
                cur.execute("SELECT data_json FROM resume_drafts WHERE id=%s", (e["draft_id"],))
                row = cur.fetchone()
                if not row:
                    print_line("  [草稿 %s] 库中已不存在，跳过" % e["draft_id"])
                    continue
                cur.execute("UPDATE resume_drafts SET data_json=%s WHERE id=%s",
                            (raw, e["draft_id"]))
                restored += 1
                print_line("  [草稿 %s] 已还原：%d 字节（sha256 %s...）" % (
                    e["draft_id"], len(raw), payload.get("sha256", "")[:12]))
    print_line("还原完成：%d / %d" % (restored, len(entries)))
    return restored > 0


def build_parser():
    p = argparse.ArgumentParser(
        prog="migrate_legacy_drafts.py",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description=(
            "存量草稿降级工具：把绑定真实模板的富文本草稿降级为纯经典草稿。\n"
            "编辑器已移除真实 Word 模板，只保留内置模板；真实模板仅用于首页展示与下载。\n"
            "本脚本不删除草稿、不删后端接口，只从 data_json 摘掉旧字段，可随时回滚。\n\n"
            "安全约定：默认只读；--apply 必须配 --yes；写库前自动备份原始 JSON。"),
        epilog=(
            "示例：\n"
            "  python tools/migrate_legacy_drafts.py\n"
            "      默认 --dry-run：只读扫描并打印诊断报告，不写库\n"
            "  python tools/migrate_legacy_drafts.py --backup\n"
            "      把受影响草稿的原始 data_json 导出到 D:/tmp/legacy_drafts_backup/\n"
            "  python tools/migrate_legacy_drafts.py --apply --yes\n"
            "      真正降级（先自动备份，再在事务里逐条移除旧字段）\n"
            "  python tools/migrate_legacy_drafts.py --rollback D:/tmp/legacy_drafts_backup --yes\n"
            "      按备份还原 data_json\n"
            "  python tools/migrate_legacy_drafts.py --only-ids 188,190,196\n"
            "      只处理指定草稿 id\n"),
    )
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", action="store_true",
                      help="只读扫描并打印报告（默认行为，不需要任何其它开关）")
    mode.add_argument("--backup", action="store_true",
                      help="把受影响草稿的原始 data_json 导出到备份目录")
    mode.add_argument("--apply", action="store_true",
                      help="把 richtext_mode/richtext_edits/template_id 从 data_json 中移除"
                           "（必须同时给 --yes 才会真正写库）")
    mode.add_argument("--rollback", metavar="备份目录",
                      help="从指定备份目录还原 data_json（必须同时给 --yes）")
    p.add_argument("--yes", action="store_true",
                   help="确认执行写库操作（--apply / --rollback 必需）")
    p.add_argument("--backup-dir", default=DEFAULT_BACKUP_DIR,
                   help="备份目录，默认 %s" % DEFAULT_BACKUP_DIR)
    p.add_argument("--only-ids", default="",
                   help="只处理指定草稿 id，逗号分隔，如 188,190,196（默认全部）")
    return p


def main(argv=None):
    _setup_stdout()
    args = build_parser().parse_args(argv)
    only_ids = [int(x) for x in args.only_ids.replace("，", ",").split(",") if x.strip()]

    # 写库前先守住开关：没有 --yes 一律不执行
    if (args.apply or args.rollback) and not args.yes:
        print_line("拒绝执行：--apply / --rollback 会写库，必须同时给 --yes。")
        print_line("当前只做只读诊断；确认要改数据请加 --yes，或先用 --backup 导出备份。")
        return 2

    if args.rollback:
        print_line("=" * 74)
        print_line("从备份还原：%s" % args.rollback)
        print_line("=" * 74)
        return 0 if rollback(args.rollback) else 1

    rows = fetch_drafts(only_ids)
    if not rows:
        print_line("未找到草稿%s。" % ("（ids=%s）" % only_ids if only_ids else ""))
        return 0
    items = analyze(rows)

    if args.apply:
        affected = [it for it in items if it["drop"]]
        print_line("=" * 74)
        print_line("降级执行 · %s" % _now())
        print_line("数据库：%s @ %s:%s" % (DB_NAME, DB_HOST, DB_PORT))
        print_line("=" * 74)
        if not affected:
            print_line("没有可降级的草稿（data_json 中已无旧字段）。")
            return 0
        # 写库前强制备份，避免误操作不可恢复
        manifest_path, paths = write_backup(affected, args.backup_dir)
        print_line("已自动备份原始 data_json 到：%s" % args.backup_dir)
        print_line("  清单：%s" % manifest_path)
        for p_ in paths:
            print_line("  - %s" % p_)
        changed = apply_downgrade(affected)
        print_line("")
        print_line("-" * 74)
        print_line("降级完成：%d 份草稿已改为纯经典草稿" % len(changed))
        print_line("回滚命令：python tools/migrate_legacy_drafts.py --rollback %s --yes"
                   % args.backup_dir.replace("\\", "/"))
        print_line("-" * 74)
        return 0

    if args.backup:
        affected = [it for it in items if it["affected"]]
        if not affected:
            print_line("没有受影响的草稿，无需备份。")
            return 0
        report(items, "backup")
        manifest_path, paths = write_backup(affected, args.backup_dir)
        summary(items)
        print_line("")
        print_line("备份目录：%s" % os.path.abspath(args.backup_dir))
        print_line("清单文件：%s" % os.path.basename(manifest_path))
        for p_ in paths:
            size = os.path.getsize(p_)
            print_line("  - %s（%d 字节）" % (os.path.basename(p_), size))
        print_line("回滚命令：python tools/migrate_legacy_drafts.py --rollback %s --yes"
                   % args.backup_dir.replace("\\", "/"))
        return 0

    # 默认只读诊断
    report(items, "dry-run")
    affected = summary(items)
    print_line("")
    print_line("下一步（本次未改动任何数据）：")
    if affected:
        print_line("  备份：python tools/migrate_legacy_drafts.py --backup")
        print_line("  降级：python tools/migrate_legacy_drafts.py --apply --yes")
    else:
        print_line("  无需处理：没有绑定真实模板或残留富文本字段的草稿。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
