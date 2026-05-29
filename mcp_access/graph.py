"""
graph.py — Access database dependency graph builder.

Generates a vis.js-compatible graph.json describing every object in an Access
database and the edges (relationships, RecordSource, ControlSource,
SourceObject, RowSource, VBA heuristics, macro actions) that connect them.

Usage through MCP:
    access_graph  db_path="C:/path/to/db.accdb"
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .code import ac_get_code, ac_list_objects
from .constants import CTRL_TYPE, DAO_FIELD_TYPE
from .controls import _get_parsed_controls
from .core import _Session
from .helpers import split_code_behind

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_SQL_START_RE = re.compile(
    r"^\s*(SELECT|INSERT|UPDATE|DELETE|TRANSFORM|PARAMETERS|WITH)\b", re.I
)

_RECORDSOURCE_RE = re.compile(
    r"^\s+RecordSource\s*=\s*\"?(.*?)\"?\s*$", re.M
)

# VBA DoCmd / QueryDefs patterns  (case-insensitive, dot-all)
_VBA_PATTERNS: list[dict[str, str]] = [
    {"regex": r'\bDoCmd\.OpenForm\s+"((?:[^"]|"")+)"',
     "group": "form",  "label": "OpenForm",  "kind": "vba-openform"},
    {"regex": r'\bDoCmd\.OpenReport\s+"((?:[^"]|"")+)"',
     "group": "report", "label": "OpenReport", "kind": "vba-openreport"},
    {"regex": r'\bDoCmd\.OpenQuery\s+"((?:[^"]|"")+)"',
     "group": "query",  "label": "OpenQuery",  "kind": "vba-openquery"},
    {"regex": r'\bDoCmd\.OpenTable\s+"((?:[^"]|"")+)"',
     "group": "table",  "label": "OpenTable",  "kind": "vba-opentable"},
    {"regex": r'\bCurrentDb\s*\(\s*\)\s*\.\s*QueryDefs\s*\(\s*"((?:[^"]|"")+)"\s*\)',
     "group": "query",  "label": "QueryDefs", "kind": "vba-querydefs"},
    {"regex": r'\bDBEngine\s*\(\s*0\s*\)\s*\(\s*0\s*\)\s*\.\s*QueryDefs\s*\(\s*"((?:[^"]|"")+)"\s*\)',
     "group": "query",  "label": "QueryDefs", "kind": "vba-querydefs"},
    {"regex": r'\bDoCmd\.RunMacro\s+"((?:[^"]|"")+)"',
     "group": "macro",  "label": "RunMacro",  "kind": "vba-runmacro"},
]

_VBA_RUNSQL_RE = re.compile(
    r'\bDoCmd\.RunSQL\s+"((?:[^"]|"")+)"', re.I | re.S
)
_VBA_SOURCEOBJECT_RE = re.compile(
    r'\.SourceObject\s*=\s*"((?:[^"]|"")+)"', re.I | re.M
)
_VBA_STRING_LITERAL_RE = re.compile(r'"((?:[^"]|"")*)"')

# Macro action/argument patterns
_MACRO_ACTION_RE = re.compile(r'^\s*Action\s*=\s*"?([A-Za-z0-9_]+)"?\s*$')
_MACRO_ARGUMENT_RE = re.compile(r'^\s*Argument\s*=\s*(.+?)\s*$')

_MACRO_ACTIONS: dict[str, tuple[str, str, str]] = {
    "OpenForm":   ("form",   "OpenForm",   "macro-openform"),
    "OpenReport": ("report", "OpenReport", "macro-openreport"),
    "OpenQuery":  ("query",  "OpenQuery",  "macro-openquery"),
    "OpenTable":  ("table",  "OpenTable",  "macro-opentable"),
}

# Regex to extract Public Sub/Function/Property declarations from VBA code.
# Captures the procedure name.  Implicit Public (no Private keyword) counts.
_VBA_PROC_DECL_RE = re.compile(
    r"^\s*(?:Public\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))"
    r"\s+(\w+)",
    re.MULTILINE | re.IGNORECASE,
)
_VBA_PRIVATE_PROC_RE = re.compile(
    r"^\s*Private\s+(?:Sub|Function|Property\s+(?:Get|Let|Set))"
    r"\s+(\w+)",
    re.MULTILINE | re.IGNORECASE,
)

# Built-in VBA / Access function names to exclude from cross-module call
# detection.  All lowercase.
_VBA_BUILTIN_NAMES: set[str] = {
    # String functions
    "asc", "ascw", "chr", "chrw", "format", "instr", "instrb",
    "instrrev", "join", "lcase", "left", "len", "lenb", "ltrim",
    "mid", "replace", "right", "space", "split", "str", "strcomp",
    "strconv", "strreverse", "trim", "rtrim", "ucase", "val", "string",
    # Type conversion
    "cbool", "cbyte", "ccur", "cdate", "cdbl", "cdec", "cint",
    "clng", "clnglng", "clngptr", "csng", "cstr", "cvar", "cverr",
    # Type checking
    "isarray", "isdate", "isempty", "iserror", "ismissing",
    "isnull", "isnumeric", "isobject", "typename", "vartype",
    # Math
    "abs", "atn", "cos", "exp", "fix", "int", "log", "rnd",
    "round", "sgn", "sin", "sqr", "tan",
    # Date/Time
    "date", "dateadd", "datediff", "datepart", "dateserial",
    "datevalue", "day", "formatdatetime", "hour", "minute",
    "month", "monthname", "now", "second", "time", "timeserial",
    "timevalue", "timer", "weekday", "weekdayname", "year",
    # I/O
    "inputbox", "msgbox",
    # File
    "curdir", "dir", "eof", "filecopy", "filedatetime", "filelen",
    "freefile", "getattr", "loc", "lof", "setattr",
    # Array
    "array", "erase", "filter", "lbound", "ubound",
    # Interaction / System
    "appactivate", "beep", "command", "doevents", "environ",
    "sendkeys", "shell",
    # Error
    "error",
    # Object / Reference
    "callbyname", "createobject", "getobject",
    # Registry
    "deletesetting", "getsetting", "savesetting",
    # Number conversion
    "hex", "oct",
    # Miscellaneous
    "choose", "iif", "nz", "partition", "qbcolor", "randomize", "rgb",
    # Access domain aggregates
    "davg", "dcount", "dfirst", "dlast", "dlookup", "dmax", "dmin",
    "dstdev", "dstdevp", "dsum", "dvar", "dvarp",
    # Access system
    "codedb", "currentdb", "currentuser", "eval", "guidfromstring",
    "hyperlinkpart", "stringfromguid", "syscmd",
}


# ---------------------------------------------------------------------------
# GraphBuilder
# ---------------------------------------------------------------------------

class GraphBuilder:
    """Mutable accumulator for graph nodes and edges."""

    def __init__(self, field_mode: str = "referenced"):
        self.nodes: dict[str, dict] = {}          # id -> node dict
        self.edges: list[dict] = []
        self._edge_dedup: set[tuple] = set()
        self._edge_counter = 0

        # name -> list of {node_id, group, name, is_data}
        self._name_targets: dict[str, list[dict]] = defaultdict(list)
        # table_name -> {field_name: data_type_str}
        self._known_table_fields: dict[str, dict[str, str]] = {}
        # sha256 -> node_id
        self._sql_cache: dict[str, str] = {}

        self.warnings: list[dict] = []
        self.field_mode = field_mode

        # Populated during scan; sorted desc by length for matching
        self._known_data_names: list[str] = []

        # Cross-module call detection (populated by build_proc_index)
        # proc_name_lower -> list of module node_ids that define it
        self._proc_index: dict[str, list[str]] = defaultdict(list)
        # Compiled regex for matching procedure calls (built lazily)
        self._proc_call_re: re.Pattern | None = None
        # Cache of module code read during proc index building
        self._module_code_cache: dict[str, str] = {}

    # ── node / edge primitives ──────────────────────────────────────────

    def add_node(
        self,
        node_id: str,
        label: str,
        group: str,
        meta: dict | None = None,
        *,
        is_data: bool = False,
    ) -> dict:
        if node_id in self.nodes:
            existing = self.nodes[node_id]
            if meta:
                existing.setdefault("meta", {}).update(meta)
            return existing

        node = {
            "id": node_id,
            "label": label,
            "group": group,
            "meta": dict(meta) if meta else {},
        }
        self.nodes[node_id] = node

        lname = label.lower()
        entry = {"node_id": node_id, "group": group, "name": label, "is_data": is_data}
        self._name_targets[lname].append(entry)
        if is_data:
            # Also register without brackets
            bare = _strip_brackets(label).lower()
            if bare != lname:
                self._name_targets[bare].append(entry)
        return node

    def add_edge(
        self,
        from_id: str,
        to_id: str,
        label: str,
        kind: str,
        arrows: str = "to",
        meta: dict | None = None,
    ) -> None:
        key = (from_id, to_id, kind, label)
        if key in self._edge_dedup:
            return
        self._edge_dedup.add(key)
        self._edge_counter += 1
        self.edges.append({
            "id": f"e{self._edge_counter}",
            "from": from_id,
            "to": to_id,
            "label": label,
            "kind": kind,
            "arrows": arrows,
            "meta": dict(meta) if meta else {},
        })

    def add_warning(self, code: str, message: str, meta: dict | None = None) -> None:
        self.warnings.append({
            "code": code,
            "message": message,
            "meta": dict(meta) if meta else {},
        })

    # ── name resolution ─────────────────────────────────────────────────

    def _targets_for_name(
        self, name: str, *, data_only: bool = False
    ) -> list[dict]:
        if not name:
            return []
        lname = name.strip().lower()
        hits = self._name_targets.get(lname, [])
        bare = _strip_brackets(lname)
        if bare != lname:
            hits = hits or self._name_targets.get(bare, [])
        if data_only:
            hits = [h for h in hits if h["is_data"]]
        return hits

    def _node_exists(self, node_id: str) -> bool:
        return node_id in self.nodes

    def _object_id(self, group: str, name: str) -> str:
        return f"{group}:{name}"

    # ── SQL node helpers ────────────────────────────────────────────────

    def _ensure_sql_node(
        self, sql_text: str, origin: str, sql_dir: str | None = None
    ) -> str:
        """Create or reuse a SQL node; returns node_id."""
        h = _text_hash(sql_text)
        if h in self._sql_cache:
            return self._sql_cache[h]

        node_id = f"sql:{h[:20]}"
        preview = _preview(sql_text, 120)
        self.add_node(node_id, f"SQL {h[:8]}", "sql", meta={
            "origin": origin,
            "sqlHash": h,
            "sqlLength": len(sql_text),
            "preview": preview,
        })
        if sql_dir:
            _write_if_missing(os.path.join(sql_dir, f"{h}.sql"), sql_text)

        self._sql_cache[h] = node_id

        # Add reference edges from SQL node to known data names
        self._add_sql_reference_edges(sql_text, node_id, sql_dir)
        return node_id

    def _add_sql_reference_edges(
        self, sql_text: str, from_id: str, sql_dir: str | None = None
    ) -> None:
        for name in _find_referenced_data_names(sql_text, self._known_data_names):
            for t in self._targets_for_name(name, data_only=True):
                self.add_edge(
                    from_id, t["node_id"], name, "sql-reference", "to",
                    {"name": name},
                )

    # ── field node helpers ──────────────────────────────────────────────

    def _ensure_field_node(
        self,
        owner_id: str,
        owner_group: str,
        owner_name: str,
        field_name: str,
        verified: bool = False,
        data_type: str | None = None,
    ) -> str | None:
        """Create a field node (respects field_mode). Returns node_id or None."""
        if self.field_mode == "none":
            return None
        node_id = f"field:{owner_group}:{owner_name}:{field_name}"
        self.add_node(node_id, field_name, "field", meta={
            "ownerName": owner_name,
            "fieldName": field_name,
            "verified": verified,
            "dataType": data_type or "",
        })
        self.add_edge(owner_id, node_id, "field", "field-owner", "to",
                      {"owner": owner_name, "field": field_name})
        return node_id

    # ── Phase 2: scan objects ───────────────────────────────────────────

    def scan_tables(self, app: Any, db: Any) -> None:
        for td in db.TableDefs:
            name: str = td.Name
            if _is_system(name):
                continue
            node_id = self._object_id("table", name)

            connect = ""
            source_table = ""
            try:
                connect = td.Connect or ""
                source_table = td.SourceTableName or ""
            except Exception:
                pass

            field_info: dict[str, str] = {}
            field_count = 0
            try:
                for fld in td.Fields:
                    fname: str = fld.Name
                    ftype = DAO_FIELD_TYPE.get(fld.Type, str(fld.Type))
                    field_info[fname] = ftype
                    field_count += 1
            except Exception:
                self.add_warning("FieldEnumFailed",
                                 f"Could not enumerate fields for table '{name}'.",
                                 {"name": name})

            self._known_table_fields[name] = field_info

            self.add_node(node_id, name, "table", meta={
                "fieldCount": field_count,
                "connect": connect,
                "sourceTable": source_table,
            }, is_data=True)

            if self.field_mode == "all":
                for fname, ftype in field_info.items():
                    self._ensure_field_node(node_id, "table", name, fname,
                                            verified=True, data_type=ftype)

    def scan_relationships(self, db_path: str) -> None:
        from .relations import ac_list_relationships
        rels = ac_list_relationships(db_path)
        for rel in rels.get("relationships", []):
            table = rel["table"]
            foreign = rel["foreign_table"]
            t_id = self._object_id("table", table)
            f_id = self._object_id("table", foreign)
            if not (self._node_exists(t_id) and self._node_exists(f_id)):
                continue
            fields_str = ", ".join(
                f"{f['local']} <-> {f['foreign']}" for f in rel.get("fields", [])
            )
            self.add_edge(f_id, t_id, rel["name"], "relation", "none",
                          {"name": rel["name"], "fields": fields_str})

    def scan_queries(self, app: Any, db: Any) -> None:
        for qd in db.QueryDefs:
            name: str = qd.Name
            if _is_system(name):
                continue
            node_id = self._object_id("query", name)
            sql = ""
            try:
                sql = qd.SQL or ""
            except Exception:
                pass
            self.add_node(node_id, name, "query", meta={
                "sqlPreview": _preview(sql, 200),
                "sqlHash": _text_hash(sql) if sql else "",
            }, is_data=True)

    def scan_ui_objects(self, app: Any) -> None:
        for obj_type, group in [
            ("AllForms", "form"),
            ("AllReports", "report"),
            ("AllMacros", "macro"),
            ("AllModules", "module"),
        ]:
            try:
                collection = getattr(app.CurrentProject, obj_type)
                for item in collection:
                    name: str = item.Name
                    if _is_system(name):
                        continue
                    node_id = self._object_id(group, name)
                    self.add_node(node_id, name, group)
            except Exception:
                pass

    def finalize_data_names(self) -> None:
        """Build sorted known-data-names list (longest first for matching)."""
        names: set[str] = set()
        for entries in self._name_targets.values():
            for e in entries:
                if e["is_data"]:
                    names.add(e["name"])
        self._known_data_names = sorted(names, key=len, reverse=True)

    def build_proc_index(self, db_path: str) -> None:
        """Pass 1: index all Public procedures across standalone modules.

        Reads each module's VBA code via VBE (with SaveAsText fallback),
        extracts Public Sub/Function/Property declarations, and builds
        ``self._proc_index`` for cross-module call detection.
        Code is cached in ``self._module_code_cache`` to avoid re-reading
        during ``analyze_module_code()``.
        """
        from .vbe import _get_code_module, _cm_all_code

        app = _Session.connect(db_path)
        module_names: list[str] = []
        try:
            for item in app.CurrentProject.AllModules:
                name: str = item.Name
                if not _is_system(name):
                    module_names.append(name)
        except Exception:
            return

        for mod_name in module_names:
            node_id = self._object_id("module", mod_name)
            if not self._node_exists(node_id):
                continue

            # Read module code (cache for later heuristic analysis)
            code = ""
            try:
                cm = _get_code_module(app, "module", mod_name)
                code = _cm_all_code(cm, f"module:{mod_name}")
            except Exception:
                try:
                    code = ac_get_code(db_path, "module", mod_name)
                except Exception:
                    continue
            if not code:
                continue

            self._module_code_cache[mod_name] = code

            # Find all Private proc names so we can exclude them
            private_names = {
                m.group(1).lower()
                for m in _VBA_PRIVATE_PROC_RE.finditer(code)
            }

            # Find all proc declarations (Public or implicit Public)
            for m in _VBA_PROC_DECL_RE.finditer(code):
                proc_name = m.group(1)
                pname_lower = proc_name.lower()
                # Skip Private procs, built-ins, and very short names
                if pname_lower in private_names:
                    continue
                if pname_lower in _VBA_BUILTIN_NAMES:
                    continue
                if len(pname_lower) < 2:
                    continue
                self._proc_index[pname_lower].append(node_id)

        # Build compiled regex for call detection
        proc_names = sorted(self._proc_index.keys(), key=len, reverse=True)
        if proc_names:
            escaped = [re.escape(n) for n in proc_names]
            alt = "|".join(escaped)
            # Match bare calls: ProcName( — but NOT object.ProcName(
            # Also match: Call ProcName
            self._proc_call_re = re.compile(
                rf"(?<![.\w])(?:{alt})\s*\("
                rf"|\bCall\s+(?:{alt})\b",
                re.IGNORECASE,
            )

    # ── Phase 3: form/report edge detection ─────────────────────────────

    def analyze_form_or_report(
        self, db_path: str, group: str, name: str, sql_dir: str | None,
        include_code: bool = True,
    ) -> None:
        object_id = self._object_id(group, name)
        try:
            export_text = ac_get_code(db_path, group, name)
        except Exception as exc:
            self.add_warning(
                "ExportFailed",
                f"Could not export {group} '{name}': {exc}",
                {"name": name, "group": group},
            )
            return

        # --- RecordSource ---
        record_source = _extract_record_source(export_text)
        rs_target = self._resolve_record_source(
            object_id, group, name, record_source, sql_dir
        )

        # --- Code behind ---
        _, vba_code = split_code_behind(export_text)

        # --- Controls ---
        try:
            parsed = _get_parsed_controls(db_path, group, name)
        except Exception:
            parsed = {"controls": []}

        for ctrl in parsed.get("controls", []):
            ctrl_name = ctrl.get("name", "")
            ctrl_type_name = ctrl.get("type_name", "")

            # SourceObject (subform/subreport)
            so = ctrl.get("source_object", "")
            if so:
                self._handle_source_object(
                    object_id, so, ctrl_name, ctrl_type_name, ctrl
                )

            # RowSource (combo/list)
            row_src = ctrl.get("row_source", "")
            if row_src:
                self._handle_row_source(
                    object_id, row_src, ctrl_name, ctrl_type_name, sql_dir
                )

            # ControlSource
            cs = ctrl.get("control_source", "")
            if cs:
                self._handle_control_source(
                    object_id, cs, ctrl_name, ctrl_type_name,
                    rs_target, sql_dir,
                )

        # --- VBA code heuristics ---
        if include_code and vba_code:
            self._analyze_code_heuristics(
                object_id, group, name, vba_code, sql_dir
            )

    def _resolve_record_source(
        self,
        owner_id: str,
        owner_group: str,
        owner_name: str,
        record_source: str | None,
        sql_dir: str | None,
    ) -> dict | None:
        """Returns target ref dict {node_id, group, name} or None."""
        if not record_source:
            return None

        # Try named target
        targets = self._targets_for_name(record_source, data_only=True)
        if targets:
            t = targets[0]
            self.add_edge(owner_id, t["node_id"], "RecordSource",
                          "recordsource", "to",
                          {"recordSource": record_source})
            return t

        # Try SQL
        if _is_likely_sql(record_source):
            sql_id = self._ensure_sql_node(
                record_source,
                f"{owner_group}:{owner_name}:RecordSource",
                sql_dir,
            )
            self.add_edge(owner_id, sql_id, "RecordSource",
                          "recordsource-sql", "to")
            return None  # no single target ref for field resolution

        self.add_warning(
            "UnresolvedRecordSource",
            f"Could not resolve RecordSource '{record_source}' on "
            f"{owner_group} '{owner_name}'.",
            {"owner": owner_name, "group": owner_group,
             "recordSource": record_source},
        )
        return None

    def _handle_source_object(
        self, owner_id: str, source_object: str,
        ctrl_name: str, ctrl_type: str, ctrl: dict,
    ) -> None:
        target_id = _resolve_source_object_target(source_object, self)
        if target_id and self._node_exists(target_id):
            meta: dict[str, Any] = {
                "controlName": ctrl_name,
                "controlType": ctrl_type,
                "sourceObject": source_object,
            }
            lmf = ctrl.get("link_master_fields", "")
            lcf = ctrl.get("link_child_fields", "")
            if lmf:
                meta["linkMasterFields"] = lmf
            if lcf:
                meta["linkChildFields"] = lcf
            self.add_edge(owner_id, target_id, "SourceObject",
                          "sourceobject", "to", meta)

    def _handle_row_source(
        self, owner_id: str, row_source: str,
        ctrl_name: str, ctrl_type: str,
        sql_dir: str | None,
    ) -> None:
        targets = self._targets_for_name(row_source, data_only=True)
        if targets:
            self.add_edge(
                owner_id, targets[0]["node_id"], "RowSource",
                "rowsource", "to",
                {"controlName": ctrl_name, "controlType": ctrl_type,
                 "rowSource": row_source},
            )
            return
        if _is_likely_sql(row_source):
            sql_id = self._ensure_sql_node(
                row_source, f"RowSource:{ctrl_name}", sql_dir
            )
            self.add_edge(
                owner_id, sql_id, "RowSource", "rowsource", "to",
                {"controlName": ctrl_name, "controlType": ctrl_type},
            )

    def _handle_control_source(
        self,
        owner_id: str,
        control_source: str,
        ctrl_name: str,
        ctrl_type: str,
        rs_target: dict | None,
        sql_dir: str | None,
    ) -> None:
        field_name = _field_from_control_source(control_source)

        if field_name and rs_target:
            owner_group = rs_target["group"]
            owner_name = rs_target["name"]
            target_node = rs_target["node_id"]

            verified = False
            data_type: str | None = None
            if owner_group == "table":
                fields = self._known_table_fields.get(owner_name, {})
                if field_name in fields:
                    verified = True
                    data_type = fields[field_name]

            fid = self._ensure_field_node(
                target_node, owner_group, owner_name,
                field_name, verified, data_type,
            )
            if fid:
                self.add_edge(owner_id, fid, "ControlSource",
                              "controlsource", "to",
                              {"controlName": ctrl_name,
                               "controlType": ctrl_type,
                               "controlSource": control_source})
        elif rs_target:
            # Expression-based control (starts with = or has operators)
            self.add_edge(
                owner_id, rs_target["node_id"], "ControlExpr",
                "control-expression", "to",
                {"controlName": ctrl_name, "controlType": ctrl_type,
                 "controlSource": control_source},
            )

    # ── Phase 4: VBA code heuristics ────────────────────────────────────

    def _analyze_code_heuristics(
        self, owner_id: str, owner_group: str, owner_name: str,
        code: str, sql_dir: str | None,
    ) -> None:
        if not code:
            return

        # DoCmd / QueryDefs patterns
        for pat in _VBA_PATTERNS:
            for m in re.finditer(pat["regex"], code, re.I | re.S):
                ref_name = m.group(1).replace('""', '"')
                if not ref_name:
                    continue
                target_id = self._object_id(pat["group"], ref_name)
                if self._node_exists(target_id):
                    self.add_edge(owner_id, target_id, pat["label"],
                                  pat["kind"], "to", {"name": ref_name})

        # DoCmd.RunSQL
        for m in _VBA_RUNSQL_RE.finditer(code):
            sql_text = m.group(1).replace('""', '"')
            if not sql_text:
                continue
            sql_id = self._ensure_sql_node(
                sql_text,
                f"{owner_group}:{owner_name}:VBA",
                sql_dir,
            )
            self.add_edge(owner_id, sql_id, "RunSQL", "vba-runsql", "to",
                          {"preview": _preview(sql_text, 80)})

        # SourceObject assignment in VBA
        for m in _VBA_SOURCEOBJECT_RE.finditer(code):
            so_value = m.group(1).replace('""', '"').strip()
            if not so_value:
                continue
            target_id = _resolve_source_object_target(so_value, self)
            if target_id and self._node_exists(target_id):
                self.add_edge(owner_id, target_id, "SourceObject",
                              "vba-sourceobject", "to",
                              {"sourceObject": so_value})

        # Type dependencies (As ClassName, New ClassName, ClassName.)
        seen_type: set[str] = set()
        for lname, entries in self._name_targets.items():
            for entry in entries:
                if entry["group"] != "module":
                    continue
                if entry["node_id"] == owner_id:
                    continue
                tgt_name = entry["name"]
                escaped = re.escape(tgt_name)
                type_pat = (
                    rf"(?:\bAs\s+{escaped}\b"
                    rf"|\bNew\s+{escaped}\b"
                    rf"|\b{escaped}\s*\.)"
                )
                if re.search(type_pat, code, re.I | re.M):
                    edge_key = f"{owner_id}->{entry['node_id']}"
                    if edge_key not in seen_type:
                        seen_type.add(edge_key)
                        self.add_edge(
                            owner_id, entry["node_id"], "uses type",
                            "vba-type-ref", "to", {"name": tgt_name},
                        )

        # Cross-module procedure calls (bare FuncName( or Call SubName)
        if self._proc_call_re:
            seen_call: set[str] = set()
            for m in self._proc_call_re.finditer(code):
                matched = m.group(0)
                # Extract the procedure name from the match
                # Strip leading 'Call ' if present, trailing '(' or whitespace
                proc_name = re.sub(
                    r"^\s*Call\s+", "", matched, flags=re.I
                ).rstrip("( \t")
                pname_lower = proc_name.lower()
                target_ids = self._proc_index.get(pname_lower, [])
                for tid in target_ids:
                    if tid == owner_id:
                        continue  # skip self-edges
                    edge_key = f"{owner_id}->{tid}:call:{pname_lower}"
                    if edge_key not in seen_call:
                        seen_call.add(edge_key)
                        self.add_edge(
                            owner_id, tid, "calls",
                            "vba-call", "to", {"procedure": proc_name},
                        )

        # Data references in string literals
        if self._known_data_names:
            literals = _VBA_STRING_LITERAL_RE.findall(code)
            if literals:
                combined = " ".join(s.replace('""', '"') for s in literals)
                seen_data: set[str] = set()
                for dname in _find_referenced_data_names(
                    combined, self._known_data_names
                ):
                    for t in self._targets_for_name(dname, data_only=True):
                        edge_key = f"{owner_id}->{t['node_id']}"
                        if edge_key not in seen_data:
                            seen_data.add(edge_key)
                            self.add_edge(
                                owner_id, t["node_id"], "uses data",
                                "vba-data-ref", "to", {"name": dname},
                            )

    # ── Phase 5: query & macro edges ────────────────────────────────────

    def analyze_query_edges(self, app: Any, db: Any, sql_dir: str | None) -> None:
        for qd in db.QueryDefs:
            name: str = qd.Name
            if _is_system(name):
                continue
            node_id = self._object_id("query", name)
            if not self._node_exists(node_id):
                continue
            sql = ""
            try:
                sql = qd.SQL or ""
            except Exception:
                continue
            if not sql:
                continue
            for dname in _find_referenced_data_names(sql, self._known_data_names):
                for t in self._targets_for_name(dname, data_only=True):
                    if t["node_id"] != node_id:
                        self.add_edge(
                            node_id, t["node_id"], dname,
                            "query-sql-reference", "to", {"name": dname},
                        )

    def analyze_macro(
        self, db_path: str, macro_name: str, sql_dir: str | None
    ) -> None:
        macro_id = self._object_id("macro", macro_name)
        if not self._node_exists(macro_id):
            return
        try:
            text = ac_get_code(db_path, "macro", macro_name)
        except Exception:
            return
        lines = text.splitlines()
        i = 0
        while i < len(lines):
            m_act = _MACRO_ACTION_RE.match(lines[i])
            if not m_act:
                i += 1
                continue
            action = m_act.group(1)
            arg_value: str | None = None
            for j in range(i + 1, min(i + 8, len(lines))):
                if _MACRO_ACTION_RE.match(lines[j]):
                    break
                m_arg = _MACRO_ARGUMENT_RE.match(lines[j])
                if m_arg:
                    arg_value = _convert_access_literal(m_arg.group(1))
                    break

            if action in _MACRO_ACTIONS and arg_value:
                grp, lbl, knd = _MACRO_ACTIONS[action]
                target_id = self._object_id(grp, arg_value)
                if self._node_exists(target_id):
                    self.add_edge(macro_id, target_id, lbl, knd, "to",
                                  {"name": arg_value})

            if action == "RunSQL" and arg_value:
                sql_id = self._ensure_sql_node(
                    arg_value, f"macro:{macro_name}", sql_dir
                )
                self.add_edge(macro_id, sql_id, "RunSQL", "macro-runsql",
                              "to", {"preview": _preview(arg_value, 80)})

            i += 1

    def analyze_module_code(
        self, db_path: str, module_name: str, sql_dir: str | None
    ) -> None:
        """Analyze a standalone module's VBA code for heuristic edges."""
        node_id = self._object_id("module", module_name)
        if not self._node_exists(node_id):
            return

        # Use cached code from build_proc_index() if available
        code = self._module_code_cache.get(module_name, "")
        if not code:
            try:
                from .vbe import _get_code_module, _cm_all_code
                app = _Session.connect(db_path)
                cm = _get_code_module(app, "module", module_name)
                code = _cm_all_code(cm, f"module:{module_name}")
            except Exception:
                try:
                    code = ac_get_code(db_path, "module", module_name)
                except Exception:
                    return
        if code:
            self._analyze_code_heuristics(
                node_id, "module", module_name, code, sql_dir
            )

    # ── Phase 6: output ─────────────────────────────────────────────────

    def build_output(
        self, db_path: str, out_dir: str, field_mode: str,
        embed_viewer: bool = True,
    ) -> dict:
        stats = self._compute_stats()
        graph = {
            "meta": {
                "database": db_path,
                "generatedAt": datetime.now(timezone.utc).isoformat(),
                "fieldNodeMode": {
                    "none": "None",
                    "referenced": "ReferencedOnly",
                    "all": "AllTableFields",
                }.get(field_mode, field_mode),
                "stats": stats,
                "warnings": self.warnings,
            },
            "nodes": list(self.nodes.values()),
            "edges": self.edges,
        }

        os.makedirs(out_dir, exist_ok=True)
        graph_path = os.path.join(out_dir, "graph.json")
        with open(graph_path, "w", encoding="utf-8") as f:
            json.dump(graph, f, ensure_ascii=False, indent=2, default=str)

        viewer_path: str | None = None
        if embed_viewer:
            viewer_path = self._write_viewer(out_dir, graph)

        return {
            "graph_path": graph_path,
            "viewer_path": viewer_path,
            "stats": stats,
            "warning_count": len(self.warnings),
        }

    def _compute_stats(self) -> dict:
        groups: dict[str, int] = {}
        for n in self.nodes.values():
            g = n["group"]
            groups[g] = groups.get(g, 0) + 1
        return {
            "nodeCount": len(self.nodes),
            "edgeCount": len(self.edges),
            "tables": groups.get("table", 0),
            "queries": groups.get("query", 0),
            "forms": groups.get("form", 0),
            "reports": groups.get("report", 0),
            "macros": groups.get("macro", 0),
            "modules": groups.get("module", 0),
            "sqlNodes": groups.get("sql", 0),
            "fieldNodes": groups.get("field", 0),
            "warnings": len(self.warnings),
        }

    def _write_viewer(self, out_dir: str, graph: dict) -> str | None:
        viewer_src = Path(__file__).parent / "viewer.html"
        if not viewer_src.exists():
            return None
        template = viewer_src.read_text(encoding="utf-8")
        graph_json = json.dumps(graph, ensure_ascii=False, default=str)
        embed_script = f"\n<script>var EMBEDDED_GRAPH = {graph_json};</script>\n"
        marker = "<!-- EMBED_GRAPH_DATA -->"
        if marker in template:
            html = template.replace(marker, embed_script)
        else:
            html = template.replace("</body>", embed_script + "</body>")
        viewer_path = os.path.join(out_dir, "index.html")
        with open(viewer_path, "w", encoding="utf-8") as f:
            f.write(html)
        return viewer_path


# ---------------------------------------------------------------------------
# Pure helpers (no COM, no side-effects)
# ---------------------------------------------------------------------------

def _is_system(name: str) -> bool:
    return name.startswith("MSys") or name.startswith("~")


def _strip_brackets(name: str) -> str:
    s = name.strip()
    if s.startswith("[") and s.endswith("]") and len(s) >= 2:
        return s[1:-1]
    return s


def _is_likely_sql(text: str) -> bool:
    if not text:
        return False
    return bool(_SQL_START_RE.match(text.strip()))


def _text_hash(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _preview(text: str, max_len: int = 120) -> str:
    if not text:
        return ""
    flat = re.sub(r"\s+", " ", text).strip()
    if len(flat) <= max_len:
        return flat
    return flat[:max_len].rstrip() + "..."


def _write_if_missing(path: str, content: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)


def _convert_access_literal(raw: str) -> str | None:
    if raw is None:
        return None
    val = raw.strip()
    if val == "Null":
        return None
    if len(val) >= 2 and val.startswith('"') and val.endswith('"'):
        return val[1:-1]
    return val


def _extract_record_source(export_text: str) -> str | None:
    """Extract RecordSource from form/report export text (before first Begin Section)."""
    cutoff = export_text.find("Begin Section")
    if cutoff < 0:
        cutoff = len(export_text)
    header = export_text[:cutoff]
    m = _RECORDSOURCE_RE.search(header)
    if m:
        val = m.group(1).strip()
        return val if val else None
    return None


def _field_from_control_source(cs: str) -> str | None:
    """Extract a simple field name from a ControlSource value.

    Returns None for expressions (starting with ``=`` or containing
    operators) — those are handled as control-expression edges instead.
    """
    if not cs:
        return None
    trimmed = cs.strip()
    if trimmed.startswith("="):
        return None
    if re.search(r"[+\-*/&()]", trimmed):
        return None
    parts = trimmed.split(".")
    candidate = parts[-1].strip()
    candidate = _strip_brackets(candidate)
    return candidate if candidate else None


def _resolve_source_object_target(
    source_object: str, builder: GraphBuilder
) -> str | None:
    """Resolve 'Form.frmName', 'Report.rptName', or bare 'frmName'."""
    so = source_object.strip()
    m = re.match(r"^(Form|Report)\.(.+)$", so, re.I)
    if m:
        grp = m.group(1).lower()
        name = m.group(2)
        return builder._object_id(grp, name)
    # Bare name — try form first, then report
    for grp in ("form", "report"):
        tid = builder._object_id(grp, so)
        if builder._node_exists(tid):
            return tid
    return None


def _find_referenced_data_names(
    text: str, known_names: list[str]
) -> list[str]:
    """Case-insensitive scan for known table/query names in text."""
    if not text or not known_names:
        return []
    hits: list[str] = []
    for name in known_names:
        escaped = re.escape(name)
        pattern = rf"(?<!\w)(?:\[{escaped}\]|{escaped})(?!\w)"
        if re.search(pattern, text, re.I):
            hits.append(name)
    return hits


# ---------------------------------------------------------------------------
# Main entry point (called from dispatcher on COM thread)
# ---------------------------------------------------------------------------

def ac_graph(
    db_path: str,
    out_dir: str | None = None,
    field_mode: str = "referenced",
    include_code_heuristics: bool = True,
    include_macro_heuristics: bool = True,
    embed_viewer: bool = True,
) -> dict:
    """Build a dependency graph for the given Access database.

    Returns a summary dict with graph_path, viewer_path, stats.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()

    abs_db = os.path.abspath(db_path)
    if out_dir is None:
        out_dir = os.path.join(os.path.dirname(abs_db), "access-graph-out")
    os.makedirs(out_dir, exist_ok=True)
    sql_dir = os.path.join(out_dir, "sql")
    os.makedirs(sql_dir, exist_ok=True)

    gb = GraphBuilder(field_mode=field_mode)

    # Phase 2: enumerate all objects → nodes
    gb.scan_tables(app, db)
    gb.scan_relationships(db_path)
    gb.scan_queries(app, db)
    gb.scan_ui_objects(app)
    gb.finalize_data_names()

    # Build cross-module procedure index (must precede code heuristics)
    if include_code_heuristics:
        gb.build_proc_index(db_path)

    # Phase 5 (queries first — order doesn't affect correctness)
    gb.analyze_query_edges(app, db, sql_dir)

    # Phase 3 + 4: form/report edges + code heuristics
    obj_list = ac_list_objects(db_path, "all")
    for form_name in obj_list.get("form", []):
        try:
            gb.analyze_form_or_report(
                db_path, "form", form_name, sql_dir,
                include_code=include_code_heuristics,
            )
        except Exception as exc:
            gb.add_warning("FormEdgeParseFailed",
                           f"Error analyzing form '{form_name}': {exc}",
                           {"name": form_name})

    for report_name in obj_list.get("report", []):
        try:
            gb.analyze_form_or_report(
                db_path, "report", report_name, sql_dir,
                include_code=include_code_heuristics,
            )
        except Exception as exc:
            gb.add_warning("ReportEdgeParseFailed",
                           f"Error analyzing report '{report_name}': {exc}",
                           {"name": report_name})

    # Phase 5: macros
    if include_macro_heuristics:
        for macro_name in obj_list.get("macro", []):
            try:
                gb.analyze_macro(db_path, macro_name, sql_dir)
            except Exception as exc:
                gb.add_warning("MacroEdgeParseFailed",
                               f"Error analyzing macro '{macro_name}': {exc}",
                               {"name": macro_name})

    # Phase 4: standalone module code
    if include_code_heuristics:
        for mod_name in obj_list.get("module", []):
            try:
                gb.analyze_module_code(db_path, mod_name, sql_dir)
            except Exception as exc:
                gb.add_warning("ModuleCodeParseFailed",
                               f"Error analyzing module '{mod_name}': {exc}",
                               {"name": mod_name})

    # Phase 6: output
    return gb.build_output(abs_db, out_dir, field_mode, embed_viewer)
