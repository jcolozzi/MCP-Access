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
from .controls import _get_parsed_controls, _scan_control_properties
from .core import _Session
from .helpers import join_wrapped_value, split_code_behind

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_SQL_START_RE = re.compile(
    r"^\s*(SELECT|INSERT|UPDATE|DELETE|TRANSFORM|PARAMETERS|WITH)\b", re.I
)

_RECORDSOURCE_RE = re.compile(r"^\s+RecordSource\s*=\s*(.*?)\s*$")

# VBA DoCmd / QueryDefs patterns  (case-insensitive, dot-all)
# Optional "(" and a leading named argument: DoCmd.OpenForm(FormName:="x")
_DOCMD_ARG = r'\s*\(?\s*(?:\w+\s*:=\s*)?"((?:[^"]|"")+)"'
# First argument as written: literal, variable or expression (classified later)
_FIRST_ARG = r'\s*\(?\s*(?:\w+\s*:=\s*)?((?:"(?:[^"]|"")*"|[^,:\n"])+)'
_VBA_PATTERNS: list[dict[str, str]] = [
    {"regex": r'\bDoCmd\.OpenForm' + _FIRST_ARG,
     "group": "form",  "label": "OpenForm",  "kind": "vba-openform"},
    {"regex": r'\bDoCmd\.OpenReport' + _FIRST_ARG,
     "group": "report", "label": "OpenReport", "kind": "vba-openreport"},
    {"regex": r'\bDoCmd\.OpenQuery' + _FIRST_ARG,
     "group": "query",  "label": "OpenQuery",  "kind": "vba-openquery"},
    {"regex": r'\bDoCmd\.OpenTable' + _FIRST_ARG,
     "group": "table",  "label": "OpenTable",  "kind": "vba-opentable"},
    {"regex": r'\.\s*QueryDefs\s*\(((?:"(?:[^"]|"")*"|[^,:\n"])+)',
     "group": "query",  "label": "QueryDefs", "kind": "vba-querydefs"},
    {"regex": r'\bDoCmd\.RunMacro' + _FIRST_ARG,
     "group": "macro",  "label": "RunMacro",  "kind": "vba-runmacro"},
]
_VBA_EVAL_RE = re.compile(r'\bEval\s*\(\s*"((?:[^"]|"")+)"', re.I)
_VBA_APP_RUN_RE = re.compile(r'\bApplication\s*\.\s*Run' + _FIRST_ARG, re.I)

_IDENT_RE = re.compile(r"[A-Za-z_]\w*")
_STRING_LITERAL_FULL_RE = re.compile(r'"((?:[^"]|"")*)"')

# Procedure boundaries and string bindings (for names held in variables)
_PROC_START_RE = re.compile(
    r"^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)*"
    r"(?:Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+\w+",
    re.I | re.M,
)
_PROC_END_RE = re.compile(r"^[ \t]*End[ \t]+(?:Sub|Function|Property)\b.*$", re.I | re.M)
_CONST_RE = re.compile(
    r'^[ \t]*(?:(Public|Global|Private|Dim)[ \t]+)?Const[ \t]+(\w+)'
    r'(?:[ \t]+As[ \t]+String)?[ \t]*=[ \t]*"((?:[^"]|"")*)"[ \t]*$',
    re.I | re.M,
)
_STR_ASSIGN_RE = re.compile(
    r'^[ \t]*(?:Let[ \t]+)?(\w+)[ \t]*=[ \t]*"((?:[^"]|"")*)"[ \t]*$', re.I | re.M
)

# Recordsets: Set rs = ...OpenRecordset(arg) / Me.RecordsetClone; With blocks
_SET_RE = re.compile(r"^[ \t]*Set[ \t]+(\w+)[ \t]*=[ \t]*(.+?)[ \t]*$", re.I | re.M)
_OPEN_RS_ARG_RE = re.compile(
    r'\.\s*OpenRecordset\s*\(\s*((?:"(?:[^"]|"")*"|[^,)\n"])+)', re.I
)
_ME_RECORDSET_RE = re.compile(r"^Me(?:\s*\.\s*Form)?\s*\.\s*Recordset(?:Clone)?$", re.I)
_WITH_BLOCK_RE = re.compile(
    r"^[ \t]*With[ \t]+(.+?)[ \t]*$(.*?)^[ \t]*End[ \t]+With\b", re.I | re.M | re.S
)
_WITH_FIELD_RE = re.compile(
    r'(?:^|(?<=[\s(=,&+\-*/<>]))(?:!(?:\[([^\]]+)\]|(\w+))'
    r'|\.\s*Fields\s*\(\s*"((?:[^"]|"")+)"\s*\))',
    re.M,
)
# Me!Field / Me("Field") / Me.Field in form or report code
_ME_FIELD_RE = re.compile(
    r'\bMe\s*(?:!\s*(?:\[([^\]]+)\]|(\w+))|\(\s*"((?:[^"]|"")+)"\s*\)'
    r'|\.\s*(\w+)\b(?!\s*\())',
    re.I,
)

# Bare identifier in SQL: not qualified, not a qualifier, not a function call
_SQL_TOKEN_RE = re.compile(
    r"(?<![.!\w\]])(?:\[([^\]]+)\]|([A-Za-z_]\w*))(?!\w|\s*[.!(])"
)

_ACCESS_LIBRARY_EXTS = {".accda", ".accdb", ".accde", ".mda", ".mdb", ".mde"}

_VBA_RUNSQL_RE = re.compile(
    r'\bDoCmd\.RunSQL' + _DOCMD_ARG, re.I | re.S
)
# db.Execute "qryX" / "UPDATE ..." and db.OpenRecordset("tbl" | "SELECT ...")
_VBA_SQL_CALL_RE = re.compile(
    r'\.\s*(Execute|OpenRecordset)\s*\(?\s*"((?:[^"]|"")+)"', re.I
)
_VBA_SOURCEOBJECT_RE = re.compile(
    r'\.SourceObject\s*=\s*"((?:[^"]|"")+)"', re.I | re.M
)
_VBA_STRING_LITERAL_RE = re.compile(r'"((?:[^"]|"")*)"')

# Forms!frm!ctl, [Forms]![frm X].[ctl], Reports!rpt — in SQL, expressions, VBA
_OBJ_BANG_REF_RE = re.compile(
    r"\[?\b(Forms|Reports)\]?\s*!\s*(?:\[([^\]]+)\]|(\w+))"
    r"(?:\s*[!.]\s*(?:\[([^\]]+)\]|(\w+)))?",
    re.I,
)
_VBA_OBJ_PAREN_REF_RE = re.compile(
    r'\b(Forms|Reports)\s*\(\s*"((?:[^"]|"")+)"\s*\)', re.I
)
# Name( inside an expression, but not obj.Name( or [x]!Name(
_EXPR_CALL_RE = re.compile(r"(?<![.\w!\]])([A-Za-z_]\w*)\s*\(")
_EVENT_PROP_RE = re.compile(r"^(?:On|Before|After)[A-Z]\w*$")
_EM_MACRO_RE = re.compile(r"^(\w+)EmMacro\s*=\s*Begin\s*$")
_BLOCK_OPEN_RE = re.compile(r"^(?:Begin\b|\w+\s*=\s*Begin\s*$)")
_NAME_PROP_RE = re.compile(r'^Name\s*=\s*"?(.*?)"?\s*$')
_VBA_DECL_LINE_RE = re.compile(
    r"^[ \t]*(?:(?:Public|Private|Friend|Static|Global)[ \t]+)*"
    r"(?:Sub|Function|Property[ \t]+(?:Get|Let|Set)|Declare)\b.*$",
    re.I | re.M,
)

# SQL: Table.Field / [Table].[Field] / alias.Field, and FROM/JOIN aliases
_SQL_NAME = r"(?:\[([^\]]+)\]|([A-Za-z_]\w*))"
_SQL_QUAL_REF_RE = re.compile(_SQL_NAME + r"\s*\.\s*(?:\[([^\]]+)\]|([A-Za-z_]\w*))")
_SQL_ALIAS_RE = re.compile(
    r"(?:\bFROM|\bJOIN|,)\s*" + _SQL_NAME
    + r"\s+(?:AS\s+)?(?:\[([^\]]+)\]|([A-Za-z_]\w*))",
    re.I,
)
_SQL_STRING_RE = re.compile(r'"[^"]*"|\'[^\']*\'')
_SQL_KEYWORDS = {
    "AS", "ON", "INNER", "LEFT", "RIGHT", "OUTER", "FULL", "CROSS", "JOIN",
    "WHERE", "GROUP", "ORDER", "HAVING", "UNION", "IN", "FROM", "SELECT",
    "SET", "INTO", "VALUES", "AND", "OR", "NOT", "BY", "WITH", "PIVOT",
    "TRANSFORM", "TOP", "DISTINCT",
}

# DAO QueryDef.Type values whose Fields describe output columns:
# dbQSelect, dbQCrosstab, dbQSetOperation (UNION)
_DAO_ROW_QUERY_TYPES = {0, 16, 128}

# Macro action/argument patterns
_MACRO_ACTION_RE = re.compile(r'^\s*Action\s*=\s*"?([A-Za-z0-9_]+)"?\s*$')
_MACRO_ARGUMENT_RE = re.compile(r'^\s*Argument\s*=\s*(.+?)\s*$')

_MACRO_ACTIONS: dict[str, tuple[str, str, str]] = {
    "OpenForm":   ("form",   "OpenForm",   "macro-openform"),
    "OpenReport": ("report", "OpenReport", "macro-openreport"),
    "OpenQuery":  ("query",  "OpenQuery",  "macro-openquery"),
    "OpenTable":  ("table",  "OpenTable",  "macro-opentable"),
    "RunMacro":   ("macro",  "RunMacro",   "macro-runmacro"),
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
        # query_name -> {output field: data_type_str}, from DAO QueryDef.Fields
        self._known_query_fields: dict[str, dict[str, str]] = {}
        # query_name -> {output field lower: (SourceTable, SourceField)}
        self._query_lineage: dict[str, dict[str, tuple[str, str]]] = {}
        # sha256 -> node_id
        self._sql_cache: dict[str, str] = {}

        self.warnings: list[dict] = []
        self._missing_seen: set[tuple] = set()
        # References whose target name is computed at runtime (see _record_dynamic)
        self.dynamic_refs: list[dict] = []
        self._dynamic_seen: set[tuple] = set()
        # Public string constants across standard modules: name -> values
        self._global_consts: dict[str, set[str]] = defaultdict(set)
        # (group, name lower) -> library node id, for objects in referenced libraries
        self._library_objects: dict[tuple[str, str], str] = {}
        # node id -> design timestamp at build time (see design_stamps)
        self.design_stamps: dict[str, str] = {}
        self.field_mode = field_mode

        # Raw SaveAsText export mode: "none" (compute rawHash/rawSize only)
        # or "debug" (also keep raw exports under <out>/raw/<group>/).
        self.raw_export_mode = "none"
        self.raw_dir: str | None = None

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

    def _warn_missing(
        self, owner_id: str, target_group: str, target: str, via: str
    ) -> None:
        """Record a literal reference to an object that is not in the database."""
        # Macro arguments may be expressions resolved at runtime.
        if target.startswith("="):
            return
        key = (owner_id, target_group, target.lower(), via)
        if key in self._missing_seen:
            return
        self._missing_seen.add(key)
        owner = self.nodes.get(owner_id, {})
        owner_group = owner.get("group", "")
        owner_name = owner.get("label", owner_id)
        self.add_warning(
            "MissingReference",
            f"{owner_group} '{owner_name}' references {target_group} "
            f"'{target}' ({via}), which does not exist in this database.",
            {"owner": owner_name, "group": owner_group, "ownerId": owner_id,
             "targetGroup": target_group, "target": target, "via": via},
        )

    def _set_raw_meta(
        self, node_id: str, text: str, group: str, name: str
    ) -> None:
        """Record rawHash/rawSize on a node from its SaveAsText export.

        ``rawSize`` is the UTF-8 byte length of the export text Python already
        reads — an approximate *logical* size (Access print sections PrtMip/
        PrtDevMode are stripped by ac_get_code), suitable for the relative
        "Complexity Hotspots" ranking. In ``raw_export_mode="debug"`` the raw
        text is also kept under ``<out>/raw/<group>/<name>.txt`` and its path
        recorded as ``rawPath``.
        """
        node = self.nodes.get(node_id)
        if node is None:
            return
        raw_path: str | None = None
        if self.raw_export_mode == "debug" and self.raw_dir:
            group_dir = os.path.join(self.raw_dir, _RAW_SUBDIR.get(group, group))
            try:
                os.makedirs(group_dir, exist_ok=True)
                raw_path = os.path.join(group_dir, f"{_safe_filename(name)}.txt")
                with open(raw_path, "w", encoding="utf-8") as f:
                    f.write(text)
            except Exception:
                raw_path = None
        node["meta"].update({
            "rawHash": _text_hash(text),
            "rawSize": len(text.encode("utf-8")),
            "rawPath": raw_path,
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

    def _resolve_named(self, group: str, name: str) -> str | None:
        """Node id for a literal object name, or None if it does not exist."""
        node_id = self._object_id(group, name)
        if self._node_exists(node_id):
            return node_id
        # RunMacro "mcrGroup.SubMacro" names a submacro inside mcrGroup.
        if group == "macro" and "." in name:
            node_id = self._object_id(group, name.split(".", 1)[0])
            if self._node_exists(node_id):
                return node_id
        return self._library_objects.get((group, name.lower()))

    def _object_id(self, group: str, name: str) -> str:
        return f"{group}:{name}"

    # ── SQL node helpers ────────────────────────────────────────────────

    def _ensure_sql_node(
        self, sql_text: str, origin: str, sql_dir: str | None = None
    ) -> str:
        """Create or reuse a SQL node; returns node_id.

        On a cache hit the new ``origin`` is appended so a deduplicated SQL
        node tracks every place the identical statement appears (drives the
        "Duplicate Inline SQL" report). ``origin`` is therefore a string for
        single-origin nodes and a list once two or more origins are merged.
        """
        h = _text_hash(sql_text)
        if h in self._sql_cache:
            cached_id = self._sql_cache[h]
            self._merge_sql_origin(cached_id, origin)
            return cached_id

        node_id = f"sql:{h[:20]}"
        preview = _preview(sql_text, 120)
        sql_path = os.path.join(sql_dir, f"{h}.sql") if sql_dir else None
        self.add_node(node_id, f"SQL {h[:8]}", "sql", meta={
            "origin": origin,
            "sqlHash": h,
            "sqlPath": sql_path,
            "sqlLength": len(sql_text),
            "preview": preview,
        })
        if sql_dir:
            _write_if_missing(os.path.join(sql_dir, f"{h}.sql"), sql_text)

        self._sql_cache[h] = node_id

        # Add reference edges from SQL node to known data names
        self._add_sql_reference_edges(sql_text, node_id, sql_dir)
        self._add_field_ref_edges(node_id, sql_text, "sql-field")
        self._link_object_refs(node_id, sql_text, {"via": "SQL"})
        return node_id

    def _merge_sql_origin(self, node_id: str, origin: str) -> None:
        """Append ``origin`` to a deduplicated SQL node's origin (str -> list)."""
        node = self.nodes.get(node_id)
        if node is None:
            return
        current = node["meta"].get("origin")
        if isinstance(current, list):
            if origin not in current:
                current.append(origin)
        elif current is None:
            node["meta"]["origin"] = origin
        elif current != origin:
            node["meta"]["origin"] = [current, origin]

    def _add_sql_reference_edges(
        self, sql_text: str, from_id: str, sql_dir: str | None = None
    ) -> None:
        for name in _find_referenced_data_names(sql_text, self._known_data_names):
            for t in self._targets_for_name(name, data_only=True):
                self.add_edge(
                    from_id, t["node_id"], name, "sql-reference", "to",
                    {"name": name},
                )

    def _add_field_ref_edges(self, from_id: str, sql: str, kind: str) -> None:
        """Edges to the fields a SQL statement uses, qualified or not.

        Sources are the tables/queries the statement names whose fields are
        known. An unqualified name is attributed only when exactly one of
        them defines it.
        """
        sources: dict[str, dict[str, str]] = {}
        groups: dict[str, str] = {}
        for name in _find_referenced_data_names(sql, self._known_data_names):
            for t in self._targets_for_name(name, data_only=True):
                if t["node_id"] == from_id:
                    continue
                fields = self._fields_of(t["group"], t["name"])
                if fields:
                    sources[t["name"]] = fields
                    groups[t["name"]] = t["group"]
                break
        pairs = _qualified_field_refs(sql, sources)
        pairs += [p for p in _bare_field_refs(sql, sources) if p not in pairs]
        for source, field in pairs:
            group = groups[source]
            fid = self._ensure_field_node(
                self._object_id(group, source), group, source, field,
                True, sources[source][field],
            )
            if fid:
                self.add_edge(from_id, fid, field, kind, "to",
                              {"field": f"{source}.{field}"})

    def _handle_named_arg(
        self, owner_id: str, pat: dict, raw_arg: str, scope: _VbaScope, pos: int,
    ) -> None:
        """Resolve the object-name argument of a DoCmd/QueryDefs call."""
        kind, value = _classify_arg(_clean_arg(raw_arg))
        via = None
        if kind == "literal":
            values = [value]
        elif kind == "variable" and value.lower() in scope.bindings_at(pos):
            values = sorted(scope.bindings_at(pos)[value.lower()])
            via = value
        else:
            if value:
                self._record_dynamic(owner_id, pat["group"], pat["label"], value)
            return
        for name in values:
            if not name:
                continue
            target_id = self._resolve_named(pat["group"], name)
            meta = {"name": name, **({"via": via} if via else {})}
            if target_id:
                self.add_edge(owner_id, target_id, pat["label"], pat["kind"],
                              "to", meta)
            else:
                self._warn_missing(
                    owner_id, pat["group"], name,
                    f"VBA {pat['label']}" + (f" via {via}" if via else ""))

    def _handle_app_run(
        self, owner_id: str, raw_arg: str, scope: _VbaScope, pos: int,
    ) -> None:
        kind, value = _classify_arg(_clean_arg(raw_arg))
        if kind == "variable" and value.lower() in scope.bindings_at(pos):
            values = sorted(scope.bindings_at(pos)[value.lower()])
        elif kind == "literal":
            values = [value]
        else:
            if value:
                self._record_dynamic(owner_id, "module", "Application.Run", value)
            return
        for full in values:
            # "Proc", "Module.Proc" or "Library.Proc"
            proc = full.rsplit(".", 1)[-1]
            for tid in self._proc_index.get(proc.lower(), []):
                if tid != owner_id:
                    self.add_edge(owner_id, tid, f"calls {proc}", "vba-call",
                                  "to", {"procedure": proc, "via": "Application.Run"})

    def _record_dynamic(
        self, owner_id: str, group: str, call: str, expr: str
    ) -> None:
        """Remember a reference whose target name is only known at runtime."""
        key = (owner_id, group, call, expr)
        if key in self._dynamic_seen:
            return
        self._dynamic_seen.add(key)
        owner = self.nodes.get(owner_id, {})
        self.dynamic_refs.append({
            "ownerId": owner_id, "owner": owner.get("label", owner_id),
            "group": group, "call": call, "expr": expr[:120],
        })

    # ── recordset / Me field references ─────────────────────────────────

    def _source_for(self, target: dict | None) -> tuple[str, str, dict] | None:
        """(group, name, fields) for a table/query/single-table-SQL target."""
        if not target:
            return None
        group, name = target["group"], target["name"]
        if group == "sql":
            if not target.get("table"):
                return None
            group, name = "table", target["table"]
        fields = self._fields_of(group, name)
        return (group, name, fields) if fields else None

    def _source_from_arg(
        self, raw_arg: str, bindings: dict[str, set[str]]
    ) -> tuple[str, str, dict] | None:
        """Recordset source from an OpenRecordset argument (name or SQL)."""
        kind, value = _classify_arg(_clean_arg(raw_arg))
        if kind == "variable":
            vals = bindings.get(value.lower(), set())
            if len(vals) != 1:
                return None
            value = next(iter(vals))
        elif kind != "literal":
            return None
        if _is_likely_sql(value):
            names = _find_referenced_data_names(value, self._known_data_names)
            if len(names) != 1:
                return None
            value = names[0]
        for t in self._targets_for_name(value, data_only=True):
            return self._source_for(t)
        return None

    def _recordset_source(
        self, expr: str, bindings: dict[str, set[str]], me_source: dict | None,
    ) -> tuple[str, str, dict] | None:
        expr = expr.strip()
        if _ME_RECORDSET_RE.match(expr):
            return self._source_for(me_source)
        m = _OPEN_RS_ARG_RE.search(expr)
        return self._source_from_arg(m.group(1), bindings) if m else None

    def _link_recordset_fields(
        self, owner_id: str, code: str, scope: _VbaScope, me_source: dict | None,
    ) -> None:
        for start, end in scope.spans:
            body = code[start:end]
            bindings = scope.bindings_at(start)
            rs_sources: dict[str, tuple | None] = {}
            for m in _SET_RE.finditer(body):
                var = m.group(1).lower()
                src = self._recordset_source(m.group(2), bindings, me_source)
                if var in rs_sources and rs_sources[var] != src:
                    rs_sources[var] = None  # re-pointed: ambiguous, skip it
                elif src is not None or var not in rs_sources:
                    rs_sources[var] = src
            for var, src in rs_sources.items():
                if not src:
                    continue
                v = re.escape(var)
                for m in re.finditer(
                    rf'\b{v}\s*(?:!\s*(?:\[([^\]]+)\]|(\w+))'
                    rf'|(?:\.\s*Fields)?\s*\(\s*"((?:[^"]|"")+)"\s*\))',
                    body, re.I,
                ):
                    fname = m.group(1) or m.group(2) or m.group(3)
                    self._link_vba_field(owner_id, src, fname, f"{var}!{fname}")
            for m in _WITH_BLOCK_RE.finditer(body):
                target = m.group(1).strip()
                src = rs_sources.get(target.lower()) or self._recordset_source(
                    target, bindings, me_source)
                if not src:
                    continue
                for f in _WITH_FIELD_RE.finditer(m.group(2)):
                    fname = f.group(1) or f.group(2) or f.group(3)
                    self._link_vba_field(owner_id, src, fname,
                                         f"With {target}: !{fname}")

    def _link_vba_field(
        self, owner_id: str, src: tuple, fname: str, via: str, *, warn: bool = True,
    ) -> bool:
        group, name, fields = src
        hit = _lookup_field(fields, fname)
        if not hit:
            if warn:
                owner = self.nodes.get(owner_id, {})
                self.add_warning(
                    "MissingField",
                    f"{owner.get('group', '')} '{owner.get('label', owner_id)}' "
                    f"reads '{fname}' ({via}), which is not a field of "
                    f"{group} '{name}'.",
                    {"owner": owner.get("label", owner_id),
                     "group": owner.get("group", ""), "ownerId": owner_id,
                     "targetGroup": group, "target": name, "field": fname,
                     "via": via},
                )
            return False
        fid = self._ensure_field_node(self._object_id(group, name), group, name,
                                      hit[0], True, hit[1])
        if fid:
            self.add_edge(owner_id, fid, hit[0], "vba-field", "to", {"via": via})
        return True

    def _link_me_fields(
        self, owner_id: str, code: str, me_source: dict, controls: frozenset,
    ) -> None:
        """Me!X / Me("X") / Me.X naming a RecordSource field that is not a control."""
        src = self._source_for(me_source)
        if not src:
            return
        for m in _ME_FIELD_RE.finditer(code):
            fname = m.group(1) or m.group(2) or m.group(3) or m.group(4)
            if fname.lower() in controls:
                continue
            # Me.X is usually a form property or method: link only real fields.
            self._link_vba_field(owner_id, src, fname, f"Me!{fname}", warn=False)

    def _link_object_refs(
        self, owner_id: str, text: str, meta: dict | None = None,
        *, vba: bool = False,
    ) -> None:
        """Edges for Forms!frm!ctl / Reports!rpt (and Forms("x") in VBA)."""
        refs: list[tuple[str, str, str | None]] = []
        for m in _OBJ_BANG_REF_RE.finditer(text):
            ctl = m.group(4) or m.group(5)
            if ctl and ctl.lower() in ("form", "report"):
                ctl = None  # Forms!frmMain.Form!sub — the subform property
            refs.append((m.group(1), m.group(2) or m.group(3), ctl))
        if vba:
            for m in _VBA_OBJ_PAREN_REF_RE.finditer(text):
                refs.append((m.group(1), m.group(2).replace('""', '"'), None))
        for coll, name, ctl in refs:
            group = "form" if coll.lower() == "forms" else "report"
            target_id = self._resolve_named(group, name)
            if target_id == owner_id:
                continue
            if not target_id:
                self._warn_missing(owner_id, group, name, f"{coll}! reference")
                continue
            edge_meta = {**(meta or {}), "name": name}
            if ctl:
                edge_meta["targetControl"] = ctl
            self.add_edge(owner_id, target_id, f"{coll.capitalize()}!{ctl or ''}",
                          "form-reference", "to", edge_meta)

    def _link_function_calls(
        self, owner_id: str, expr: str, label: str, kind: str, meta: dict,
    ) -> None:
        """Edges to modules defining the public functions called in an expression."""
        if not self._proc_index:
            return
        for m in _EXPR_CALL_RE.finditer(_blank_string_literals(expr)):
            name = m.group(1)
            low = name.lower()
            if low in _VBA_BUILTIN_NAMES:
                continue
            for tid in self._proc_index.get(low, []):
                if tid != owner_id:
                    self.add_edge(owner_id, tid, label, kind, "to",
                                  {**meta, "procedure": name})

    def _fields_of(self, group: str, name: str) -> dict[str, str] | None:
        """Known fields of a table/query, or None when they are unknown."""
        if group == "table":
            return self._known_table_fields.get(name)
        if group == "query":
            return self._known_query_fields.get(name)
        return None

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
        existing = self.nodes.get(node_id)
        if existing is not None:
            # Never downgrade: a later unverified reference must not unset it.
            meta = existing["meta"]
            meta["verified"] = bool(meta.get("verified")) or verified
            if data_type and not meta.get("dataType"):
                meta["dataType"] = data_type
            return node_id
        self.add_node(node_id, field_name, "field", meta={
            "ownerId": owner_id,
            "ownerGroup": owner_group,
            "ownerName": owner_name,
            "fieldName": field_name,
            "verified": verified,
            "dataType": data_type or "",
        })
        self.add_edge(owner_id, node_id, "field", "field-owner", "to",
                      {"owner": owner_name, "field": field_name})
        if owner_group == "query":
            self._link_field_lineage(owner_name, field_name)
        return node_id

    def set_query_fields(
        self, query_name: str, fields: list[tuple[str, str, str, str]]
    ) -> None:
        """Record DAO output fields: (name, SourceTable, SourceField, data type)."""
        self._known_query_fields[query_name] = {n: dt for n, _, _, dt in fields}
        self._query_lineage[query_name] = {
            n.lower(): (st, sf) for n, st, sf, _ in fields if st and sf
        }

    def _link_field_lineage(self, query_name: str, field_name: str) -> None:
        """Link a query output field to the table/query field it comes from."""
        src = self._query_lineage.get(query_name, {}).get(field_name.lower())
        if not src:
            return
        source_name, source_field = src
        targets = self._targets_for_name(source_name, data_only=True)
        query_id = self._object_id("query", query_name)
        if not targets or targets[0]["node_id"] == query_id:
            return
        t = targets[0]
        known = self._fields_of(t["group"], t["name"])
        hit = _lookup_field(known, source_field) if known is not None else None
        src_fid = self._ensure_field_node(
            t["node_id"], t["group"], t["name"],
            hit[0] if hit else source_field, bool(hit), hit[1] if hit else None,
        )
        if src_fid:
            self.add_edge(
                f"field:query:{query_name}:{field_name}", src_fid, "from",
                "field-lineage", "to",
                {"source": f"{source_name}.{source_field}"},
            )

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
            connect = ""
            try:
                sql = qd.SQL or ""
            except Exception:
                pass
            try:
                connect = qd.Connect or ""
            except Exception:
                pass
            self.add_node(node_id, name, "query", meta={
                "connect": connect,
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
            is_class = False
            try:
                cm = _get_code_module(app, "module", mod_name)
                code = _cm_all_code(cm, f"module:{mod_name}")
                try:
                    is_class = int(cm.Parent.Type) == 2  # vbext_ct_ClassModule
                except Exception:
                    pass
            except Exception:
                try:
                    code = ac_get_code(db_path, "module", mod_name)
                    is_class = code.lstrip().upper().startswith("VERSION 1.0 CLASS")
                except Exception:
                    continue
            if not code:
                continue

            self._module_code_cache[mod_name] = code
            self.index_module_procs(node_id, code, is_class=is_class)
            for k, v in _public_consts(code).items():
                self._global_consts[k] |= v

        self._compile_proc_call_re()

    def scan_vba_projects(self, app: Any) -> None:
        """Index referenced library databases and report broken VBA references."""
        from .core import _get_vb_project
        try:
            for ref in _get_vb_project(app).References:
                try:
                    broken = bool(ref.IsBroken)
                except Exception:
                    broken = True
                if broken:
                    name = _safe_attr(ref, "Name") or _safe_attr(ref, "Guid")
                    path = _safe_attr(ref, "FullPath")
                    self.add_warning(
                        "BrokenReference",
                        f"VBA reference '{name}' is broken ({path or 'path unknown'}); "
                        "the project will not compile until it is fixed.",
                        {"owner": name, "group": "reference", "target": path},
                    )
        except Exception:
            pass

        try:
            host = str(app.CurrentProject.FullName)
            projects = list(app.VBE.VBProjects)
        except Exception:
            return
        for proj in projects:
            path = _safe_attr(proj, "FileName")
            if (not path or _same_file(path, host)
                    or os.path.splitext(path)[1].lower() not in _ACCESS_LIBRARY_EXTS):
                continue
            self.add_library(
                _safe_attr(proj, "Name") or os.path.basename(path), path,
                _library_components(proj))

    def add_library(self, name: str, path: str, components: dict | None) -> str:
        """Register a library project: its forms/reports resolve, its procs are callable.

        ``components`` is ``{"forms": [...], "reports": [...], "modules":
        {name: code}}``, or None when the project is locked (e.g. an .accde).
        """
        lib_id = f"library:{name}"
        comps = components or {"forms": [], "reports": [], "modules": {}}
        self.add_node(lib_id, name, "library", meta={
            "path": path, "locked": components is None,
            "forms": comps["forms"], "reports": comps["reports"],
            "modules": sorted(comps["modules"]),
        })
        for group, key in (("form", "forms"), ("report", "reports")):
            for obj in comps[key]:
                self._library_objects.setdefault((group, obj.lower()), lib_id)
        for mod, code in comps["modules"].items():
            self._library_objects.setdefault(("module", mod.lower()), lib_id)
            self.index_module_procs(lib_id, code)
            for k, v in _public_consts(code).items():
                self._global_consts[k] |= v
        return lib_id

    def index_module_procs(
        self, node_id: str, code: str, *, is_class: bool = False
    ) -> None:
        """Add a standard module's public procedures to the call index.

        Class modules are skipped: their methods are only reachable through an
        instance (obj.Method), which the call regex deliberately ignores.
        """
        if is_class:
            return
        private_names = {
            m.group(1).lower() for m in _VBA_PRIVATE_PROC_RE.finditer(code)
        }
        for m in _VBA_PROC_DECL_RE.finditer(code):
            pname_lower = m.group(1).lower()
            if (pname_lower in private_names
                    or pname_lower in _VBA_BUILTIN_NAMES
                    or len(pname_lower) < 2):
                continue
            if node_id not in self._proc_index[pname_lower]:
                self._proc_index[pname_lower].append(node_id)

    def _compile_proc_call_re(self) -> None:
        proc_names = sorted(self._proc_index.keys(), key=len, reverse=True)
        if not proc_names:
            self._proc_call_re = None
            return
        alt = "|".join(re.escape(n) for n in proc_names)
        # Foo(...)  |  Call Foo  |  Foo a, b  (statement start, not Foo = / Foo.x)
        self._proc_call_re = re.compile(
            rf"(?<![.\w])(?P<paren>{alt})\s*\("
            rf"|\bCall\s+(?P<call>{alt})\b"
            rf"|(?:^|:|\bThen\b|\bElse\b)[ \t]*(?P<stmt>{alt})\b(?![ \t]*[=(.!:])",
            re.IGNORECASE | re.MULTILINE,
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

        self._set_raw_meta(object_id, export_text, group, name)

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
            controls = frozenset(
                c.get("name", "").lower() for c in parsed.get("controls", []))
            self._analyze_code_heuristics(
                object_id, group, name, vba_code, sql_dir,
                me_source=rs_target, me_controls=controls,
            )

        # --- Event properties, expressions, embedded macros ---
        self._analyze_properties(object_id, export_text, sql_dir)

    def _analyze_properties(
        self, owner_id: str, export_text: str, sql_dir: str | None
    ) -> None:
        try:
            props = _scan_control_properties(export_text)
        except Exception:
            props = []
        for p in props:
            value = p["value"].strip()
            if not value:
                continue
            prop, ctrl = p["property"], p["control"]
            where = f"{ctrl}.{prop}" if ctrl else prop
            meta: dict[str, Any] = {"property": prop}
            if ctrl:
                meta["controlName"] = ctrl
            is_event = bool(_EVENT_PROP_RE.match(prop))
            if value.startswith("="):
                self._link_function_calls(
                    owner_id, value, where,
                    "event-call" if is_event else "expression-call", meta,
                )
                self._link_object_refs(owner_id, value, meta)
            elif is_event and not value.startswith("["):
                # [Event Procedure] / [Embedded Macro] start with "["; else a macro name
                target_id = self._resolve_named("macro", value)
                if target_id:
                    self.add_edge(owner_id, target_id, where, "event-macro",
                                  "to", {**meta, "macro": value})
                else:
                    self._warn_missing(owner_id, "macro", value, where)
        for block in _embedded_macro_blocks(export_text):
            where = (f"{block['control']}.{block['property']}"
                     if block["control"] else block["property"])
            self._analyze_macro_lines(owner_id, block["lines"], sql_dir,
                                      {"embedded": where})

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
            # A single-table SQL RecordSource binds fields of that table.
            names = _find_referenced_data_names(record_source,
                                                self._known_data_names)
            table = (names[0] if len(names) == 1
                     and names[0] in self._known_table_fields else None)
            return {"node_id": sql_id, "group": "sql", "name": sql_id,
                    "table": table}

        self.add_warning(
            "UnresolvedRecordSource",
            f"Could not resolve RecordSource '{record_source}' on "
            f"{owner_group} '{owner_name}'.",
            {"owner": owner_name, "group": owner_group, "ownerId": owner_id,
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
        else:
            self._warn_missing(owner_id, _source_object_group(source_object),
                               source_object, f"SourceObject of {ctrl_name}")

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
        meta = {"controlName": ctrl_name, "controlType": ctrl_type,
                "controlSource": control_source}

        if field_name and rs_target:
            target = rs_target
            if rs_target["group"] == "sql":
                if not rs_target.get("table"):
                    return
                table = rs_target["table"]
                target = {"node_id": self._object_id("table", table),
                          "group": "table", "name": table}
            known = self._fields_of(target["group"], target["name"])
            hit = _lookup_field(known, field_name) if known is not None else None
            if known is not None and not hit and "." in control_source:
                # DAO names a column ambiguous across a join "Table.Field".
                qualified = ".".join(_strip_brackets(p) for p in control_source.split("."))
                hit = _lookup_field(known, qualified)
            if known is not None and not hit:
                if rs_target["group"] == "sql":
                    return  # an alias or computed column of the inline SQL
                owner = self.nodes.get(owner_id, {})
                self.add_warning(
                    "MissingField",
                    f"{owner.get('group', '')} '{owner.get('label', owner_id)}' "
                    f"control '{ctrl_name}' is bound to '{field_name}', which is "
                    f"not a field of {target['group']} '{target['name']}'.",
                    {"owner": owner.get("label", owner_id),
                     "group": owner.get("group", ""), "ownerId": owner_id,
                     "targetGroup": target["group"], "target": target["name"],
                     "field": field_name, "via": f"ControlSource of {ctrl_name}"},
                )
            fid = self._ensure_field_node(
                target["node_id"], target["group"], target["name"],
                hit[0] if hit else field_name, bool(hit),
                hit[1] if hit else None,
            )
            if fid:
                self.add_edge(owner_id, fid, "ControlSource",
                              "controlsource", "to", meta)
        elif rs_target:
            # Expression-based control (starts with = or has operators)
            self.add_edge(
                owner_id, rs_target["node_id"], "ControlExpr",
                "control-expression", "to", meta,
            )

    # ── Phase 4: VBA code heuristics ────────────────────────────────────

    def _analyze_code_heuristics(
        self, owner_id: str, owner_group: str, owner_name: str,
        code: str, sql_dir: str | None,
        me_source: dict | None = None, me_controls: frozenset = frozenset(),
    ) -> None:
        if not code:
            return
        code = _strip_vba_comments(code)
        scope = _VbaScope(code, self._global_consts)

        # DoCmd / QueryDefs: literal, variable/constant, or runtime expression
        for pat in _VBA_PATTERNS:
            for m in re.finditer(pat["regex"], code, re.I):
                self._handle_named_arg(owner_id, pat, m.group(1),
                                       scope, m.start())

        # Eval("Fn()") and Application.Run "Proc"
        for m in _VBA_EVAL_RE.finditer(code):
            self._link_function_calls(owner_id, m.group(1).replace('""', '"'),
                                      "Eval", "vba-call", {"via": "Eval"})
        for m in _VBA_APP_RUN_RE.finditer(code):
            self._handle_app_run(owner_id, m.group(1), scope, m.start())

        # Recordset and Me field references
        self._link_recordset_fields(owner_id, code, scope, me_source)
        if me_source:
            self._link_me_fields(owner_id, code, me_source, me_controls)

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

        # db.Execute / db.OpenRecordset with a saved query/table name or SQL
        for m in _VBA_SQL_CALL_RE.finditer(code):
            method = m.group(1)
            arg = m.group(2).replace('""', '"').strip()
            if not arg:
                continue
            if _is_likely_sql(arg):
                sql_id = self._ensure_sql_node(
                    arg, f"{owner_group}:{owner_name}:VBA", sql_dir)
                self.add_edge(owner_id, sql_id, method, "vba-runsql", "to",
                              {"preview": _preview(arg, 80)})
                continue
            for t in self._targets_for_name(arg, data_only=True):
                self.add_edge(owner_id, t["node_id"], method, "vba-data-ref",
                              "to", {"name": arg})

        # Forms!frm!ctl, Forms("frm"), Reports!rpt
        self._link_object_refs(owner_id, code, {"via": "VBA"}, vba=True)

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
            else:
                self._warn_missing(owner_id, _source_object_group(so_value),
                                   so_value, "VBA SourceObject")

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

        # Cross-module procedure calls: Foo(, Call Foo, or Foo a, b
        if self._proc_call_re:
            # A same-named procedure in this module shadows the public one.
            local = {
                m.group(1).lower()
                for rx in (_VBA_PROC_DECL_RE, _VBA_PRIVATE_PROC_RE)
                for m in rx.finditer(code)
            }
            call_code = _blank_string_literals(_VBA_DECL_LINE_RE.sub("", code))
            for m in self._proc_call_re.finditer(call_code):
                proc_name = m.group("paren") or m.group("call") or m.group("stmt")
                pname_lower = proc_name.lower()
                if pname_lower in local:
                    continue
                for tid in self._proc_index.get(pname_lower, []):
                    if tid != owner_id:
                        self.add_edge(
                            owner_id, tid, f"calls {proc_name}",
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
        queries: list[tuple[str, str]] = []
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
            self._link_object_refs(node_id, sql, {"via": "SQL"})
            self._read_query_fields(qd, name, node_id)
            queries.append((node_id, sql))

        # Second pass: every query's output fields are known by now, so a
        # query built on another query can attribute its field names.
        for node_id, sql in queries:
            self._add_field_ref_edges(node_id, sql, "query-field")

        if self.field_mode == "all":
            for qname, fields in self._known_query_fields.items():
                for fname, dtype in fields.items():
                    self._ensure_field_node(self._object_id("query", qname),
                                            "query", qname, fname, True, dtype)

    def _read_query_fields(self, qd: Any, name: str, node_id: str) -> None:
        """Record a row-returning query's output fields and their DAO lineage."""
        try:
            qtype = int(qd.Type)
        except Exception:
            return
        if qtype not in _DAO_ROW_QUERY_TYPES:
            return
        fields: list[tuple[str, str, str, str]] = []
        try:
            for fld in qd.Fields:
                fields.append((
                    fld.Name,
                    _safe_attr(fld, "SourceTable"),
                    _safe_attr(fld, "SourceField"),
                    DAO_FIELD_TYPE.get(fld.Type, str(fld.Type)),
                ))
        except Exception as exc:
            self.add_warning(
                "QueryFieldsUnavailable",
                f"Could not resolve the output fields of query '{name}' "
                f"(often a missing table, query or field): {exc}",
                {"owner": name, "group": "query", "ownerId": node_id},
            )
            return
        self.set_query_fields(name, fields)

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
        self._set_raw_meta(macro_id, text, "macro", macro_name)
        self._analyze_macro_lines(macro_id, text.splitlines(), sql_dir)

    def _analyze_macro_lines(
        self, owner_id: str, lines: list[str], sql_dir: str | None,
        extra_meta: dict | None = None,
    ) -> None:
        """Edges for Action/Argument pairs — standalone or embedded macros."""
        extra = dict(extra_meta or {})
        embedded = extra.get("embedded")
        via = f"embedded macro {embedded}" if embedded else "macro"
        origin = f"{owner_id}:{embedded}" if embedded else owner_id
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
            i += 1
            if not arg_value:
                continue

            if action in _MACRO_ACTIONS:
                grp, lbl, knd = _MACRO_ACTIONS[action]
                target_id = self._resolve_named(grp, arg_value)
                if not target_id:
                    self._warn_missing(owner_id, grp, arg_value, f"{via} {lbl}")
                elif target_id != owner_id:
                    self.add_edge(owner_id, target_id, lbl, knd, "to",
                                  {**extra, "name": arg_value})
            elif action == "RunCode":
                self._link_function_calls(owner_id, arg_value, "RunCode",
                                          "macro-runcode", extra)
            elif action == "RunSQL":
                sql_id = self._ensure_sql_node(arg_value, origin, sql_dir)
                self.add_edge(owner_id, sql_id, "RunSQL", "macro-runsql", "to",
                              {**extra, "preview": _preview(arg_value, 80)})

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
            self._set_raw_meta(node_id, code, "module", module_name)
            self._analyze_code_heuristics(
                node_id, "module", module_name, code, sql_dir
            )

    def export_raw_remaining(self, db_path: str, obj_list: dict) -> None:
        """Debug-mode supplemental pass: ensure every UI/query object has raw
        meta + a kept raw file, even when its heuristic pass was skipped.

        Idempotent — objects already carrying ``rawHash`` (forms/reports, and
        macros/modules analyzed with heuristics on) are left untouched; this
        fills the gaps (queries always; macros/modules when heuristics off).
        """
        for group in ("query", "form", "report", "macro", "module"):
            for name in obj_list.get(group, []):
                node_id = self._object_id(group, name)
                node = self.nodes.get(node_id)
                if node is None or "rawHash" in node["meta"]:
                    continue
                try:
                    text = ac_get_code(db_path, group, name)
                except Exception:
                    continue
                self._set_raw_meta(node_id, text, group, name)

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
                "designStamps": self.design_stamps,
                "dynamicReferences": self.dynamic_refs,
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
            "warnings_by_code": self._warnings_by_code(),
        }

    def _warnings_by_code(self) -> dict[str, int]:
        counts: dict[str, int] = defaultdict(int)
        for w in self.warnings:
            counts[w["code"]] += 1
        return dict(counts)

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
            "libraries": groups.get("library", 0),
            "dynamicReferences": len(self.dynamic_refs),
            "warnings": len(self.warnings),
        }

    def _write_viewer(self, out_dir: str, graph: dict) -> str | None:
        viewer_src = Path(__file__).parent / "viewer.html"
        if not viewer_src.exists():
            return None
        template = viewer_src.read_text(encoding="utf-8")
        # Database text is untrusted: a literal "</script>" would end the tag.
        graph_json = json.dumps(
            graph, ensure_ascii=False, default=str
        ).replace("<", "\\u003c")
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


def _same_file(a: str, b: str) -> bool:
    """Path equality that survives 8.3 short names and mapped drive vs UNC."""
    try:
        return os.path.samefile(a, b)
    except OSError:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def _library_components(proj: Any) -> dict | None:
    """Forms/reports with a code module and standard-module code of a VBProject.

    None when the project cannot be read (locked or compiled-only).
    """
    out: dict[str, Any] = {"forms": [], "reports": [], "modules": {}}
    try:
        for comp in proj.VBComponents:
            cname = str(comp.Name)
            ctype = int(comp.Type)
            if ctype == 100 and cname.startswith("Form_"):
                out["forms"].append(cname[5:])
            elif ctype == 100 and cname.startswith("Report_"):
                out["reports"].append(cname[7:])
            elif ctype == 1:  # vbext_ct_StdModule
                cm = comp.CodeModule
                n = int(cm.CountOfLines)
                out["modules"][cname] = str(cm.Lines(1, n)) if n else ""
    except Exception:
        return None
    return out


def design_stamps(app: Any, db: Any) -> dict[str, str]:
    """node id -> last design change (DAO LastUpdated / AccessObject.DateModified).

    Unlike the file's mtime, these move only on design edits — not on data
    writes, and not on the rewrite Access does when it closes the database.
    """
    stamps: dict[str, str] = {}
    for coll, group in ((db.TableDefs, "table"), (db.QueryDefs, "query")):
        for obj in coll:
            if not _is_system(obj.Name):
                stamps[f"{group}:{obj.Name}"] = _safe_attr(obj, "LastUpdated")
    for attr, group in (("AllForms", "form"), ("AllReports", "report"),
                        ("AllMacros", "macro"), ("AllModules", "module")):
        try:
            for item in getattr(app.CurrentProject, attr):
                if not _is_system(item.Name):
                    stamps[f"{group}:{item.Name}"] = _safe_attr(item, "DateModified")
        except Exception:
            pass
    return stamps


def _safe_attr(obj: Any, name: str) -> str:
    """A COM property as text, or "" when it raises (e.g. calculated fields)."""
    try:
        value = getattr(obj, name)
    except Exception:
        return ""
    return str(value) if value is not None else ""


_REM_RE = re.compile(r"^\s*Rem(\s|$)", re.I)


def _strip_vba_comments(code: str) -> str:
    """Blank out VBA comments, keeping line count and string literals intact."""
    out: list[str] = []
    for line in code.splitlines():
        if _REM_RE.match(line):
            out.append("")
            continue
        in_str = False
        cut = len(line)
        for i, ch in enumerate(line):
            if ch == '"':
                in_str = not in_str  # "" inside a literal toggles twice
            elif ch == "'" and not in_str:
                cut = i
                break
        out.append(line[:cut])
    return "\n".join(out)


def _blank_string_literals(code: str) -> str:
    """Replace the contents of VBA string literals with spaces (quotes kept)."""
    return _VBA_STRING_LITERAL_RE.sub(
        lambda m: '"' + " " * len(m.group(1)) + '"', code
    )


def _clean_arg(raw: str) -> str:
    """Cut a captured argument at the first unbalanced ')' outside strings."""
    depth = 0
    in_str = False
    for i, ch in enumerate(raw):
        if ch == '"':
            in_str = not in_str
        elif not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth < 0:
                    return raw[:i].strip()
    return raw.strip()


def _classify_arg(arg: str) -> tuple[str, str]:
    """('literal', text) | ('variable', name) | ('expression', arg)."""
    m = _STRING_LITERAL_FULL_RE.fullmatch(arg)
    if m:
        return "literal", m.group(1).replace('""', '"')
    if _IDENT_RE.fullmatch(arg):
        return "variable", arg
    return "expression", arg


def _proc_spans(code: str) -> list[tuple[int, int]]:
    """Character spans of each procedure, from its header to its End line."""
    spans: list[tuple[int, int]] = []
    ends = [m.end() for m in _PROC_END_RE.finditer(code)]
    for m in _PROC_START_RE.finditer(code):
        end = next((e for e in ends if e > m.start()), len(code))
        if not spans or m.start() >= spans[-1][1]:
            spans.append((m.start(), end))
    return spans


def _string_bindings(text: str) -> dict[str, set[str]]:
    """name (lower) -> string literals assigned to it by Const or plain assignment."""
    out: dict[str, set[str]] = defaultdict(set)
    for m in _CONST_RE.finditer(text):
        out[m.group(2).lower()].add(m.group(3).replace('""', '"'))
    for m in _STR_ASSIGN_RE.finditer(text):
        out[m.group(1).lower()].add(m.group(2).replace('""', '"'))
    return out


def _public_consts(code: str) -> dict[str, set[str]]:
    """Module-level Public/Global string constants of a standard module."""
    spans = _proc_spans(code)
    out: dict[str, set[str]] = defaultdict(set)
    for m in _CONST_RE.finditer(code):
        if (m.group(1) or "").lower() not in ("public", "global"):
            continue
        if any(s <= m.start() < e for s, e in spans):
            continue
        out[m.group(2).lower()].add(m.group(3).replace('""', '"'))
    return out


class _VbaScope:
    """Resolves a variable to the string literals it can hold at a position.

    Scope is approximate on purpose: procedure-local bindings, then the
    module's own (module-level) bindings, then public constants of every
    standard module. Parameters and computed values stay unresolved.
    """

    def __init__(self, code: str, global_consts: dict[str, set[str]]):
        self.spans = _proc_spans(code)
        outside = code
        for s, e in reversed(self.spans):
            outside = outside[:s] + "\n" * code.count("\n", s, e) + outside[e:]
        self._module = _string_bindings(outside)
        self._procs = [_string_bindings(code[s:e]) for s, e in self.spans]
        self._global = global_consts

    def bindings_at(self, pos: int) -> dict[str, set[str]]:
        merged: dict[str, set[str]] = defaultdict(set)
        for src in (self._global, self._module):
            for k, v in src.items():
                merged[k] |= v
        for (s, e), proc in zip(self.spans, self._procs):
            if s <= pos < e:
                for k, v in proc.items():
                    merged[k] = set(v)  # a local binding shadows outer ones
                break
        return merged


def _bare_field_refs(
    sql: str, sources: dict[str, dict[str, str]]
) -> list[tuple[str, str]]:
    """(source, field) for unqualified names that exactly one source defines."""
    text = _SQL_STRING_RE.sub("''", sql)
    owners: dict[str, list[tuple[str, str]]] = defaultdict(list)
    for src, fields in sources.items():
        for f in fields:
            owners[f.lower()].append((src, f))
    source_names = {s.lower() for s in sources}
    out: list[tuple[str, str]] = []
    for m in _SQL_TOKEN_RE.finditer(text):
        tok = m.group(1) or m.group(2)
        low = tok.lower()
        if (m.group(2) and tok.upper() in _SQL_KEYWORDS) or low in source_names:
            continue
        hits = owners.get(low, [])
        if len(hits) == 1 and hits[0] not in out:
            out.append(hits[0])
    return out


def _lookup_field(fields: dict[str, str], name: str) -> tuple[str, str] | None:
    """(canonical name, data type) of a field, matched case-insensitively."""
    if name in fields:
        return name, fields[name]
    low = name.lower()
    for k, v in fields.items():
        if k.lower() == low:
            return k, v
    return None


def _qualified_field_refs(
    sql: str, table_fields: dict[str, dict[str, str]]
) -> list[tuple[str, str]]:
    """(table, field) for every Table.Field / alias.Field naming a real table field."""
    text = _SQL_STRING_RE.sub("''", sql)
    by_lower = {t.lower(): t for t in table_fields}
    aliases: dict[str, str] = {}
    for m in _SQL_ALIAS_RE.finditer(text):
        table = by_lower.get((m.group(1) or m.group(2)).lower())
        alias = m.group(3) or m.group(4)
        if table and alias.upper() not in _SQL_KEYWORDS:
            aliases[alias.lower()] = table
    out: list[tuple[str, str]] = []
    for m in _SQL_QUAL_REF_RE.finditer(text):
        qual = (m.group(1) or m.group(2)).lower()
        table = aliases.get(qual) or by_lower.get(qual)
        if not table:
            continue
        hit = _lookup_field(table_fields[table], m.group(3) or m.group(4))
        if hit and (table, hit[0]) not in out:
            out.append((table, hit[0]))
    return out


def _block_end(lines: list[str], start: int) -> int:
    """Index of the End closing the block opened on ``lines[start]``."""
    depth = 0
    for j in range(start, len(lines)):
        s = lines[j].strip()
        if _BLOCK_OPEN_RE.match(s):
            depth += 1
        elif s == "End":
            depth -= 1
            if depth == 0:
                return j
    return len(lines) - 1


def _embedded_macro_blocks(text: str) -> list[dict]:
    """``[{control, property, lines}]`` for each ``OnXxxEmMacro = Begin`` block.

    The owning block's ``Name`` may come after the macro, so it is assigned
    when that block closes. Form-level macros get ``control: ""``.
    """
    lines = text.splitlines()
    out: list[dict] = []
    stack: list[dict] = []

    def _flush(frame: dict) -> None:
        for entry in frame["pending"]:
            entry["control"] = frame["name"]
            out.append(entry)

    i = 0
    while i < len(lines):
        s = lines[i].strip()
        if s in ("CodeBehindForm", "CodeBehindReport"):
            break
        m = _EM_MACRO_RE.match(s)
        if m:
            end = _block_end(lines, i)
            entry = {"control": "", "property": m.group(1),
                     "lines": lines[i + 1:end]}
            (stack[-1]["pending"] if stack else out).append(entry)
            i = end + 1
            continue
        if _BLOCK_OPEN_RE.match(s):
            if s.startswith("Begin"):
                stack.append({"name": "", "pending": []})
                i += 1
            else:
                i = _block_end(lines, i) + 1  # other Prop = Begin blocks
            continue
        if s == "End":
            if stack:
                _flush(stack.pop())
            i += 1
            continue
        m = _NAME_PROP_RE.match(s)
        if m and stack and not stack[-1]["name"]:
            stack[-1]["name"] = m.group(1)
        i += 1
    while stack:
        _flush(stack.pop())
    return out


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


_RAW_SUBDIR = {
    "form": "forms",
    "report": "reports",
    "query": "queries",
    "macro": "macros",
    "module": "modules",
}


def _safe_filename(name: str) -> str:
    """Strip characters illegal in Windows filenames."""
    if not name:
        return "_"
    return re.sub(r'[\\/:*?"<>|]', "_", name)


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
    lines = export_text[:cutoff].splitlines()
    for i, line in enumerate(lines):
        m = _RECORDSOURCE_RE.match(line)
        if m:
            val, _ = join_wrapped_value(lines, i, m.group(1))
            val = val.strip()
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


_SOURCE_OBJECT_PREFIX_RE = re.compile(r"^(Form|Report|Table|Query)\.(.+)$", re.I)


def _source_object_group(source_object: str) -> str:
    m = _SOURCE_OBJECT_PREFIX_RE.match(source_object.strip())
    return m.group(1).lower() if m else "form"


def _resolve_source_object_target(
    source_object: str, builder: GraphBuilder
) -> str | None:
    """Resolve 'Form.x', 'Report.x', 'Table.x', 'Query.x', or a bare form/report name."""
    so = source_object.strip()
    m = _SOURCE_OBJECT_PREFIX_RE.match(so)
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
    raw_export_mode: str = "none",
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

    raw_export_mode = (raw_export_mode or "none").lower()
    if raw_export_mode not in ("none", "debug"):
        raw_export_mode = "none"

    gb = GraphBuilder(field_mode=field_mode)
    gb.raw_export_mode = raw_export_mode
    if raw_export_mode == "debug":
        gb.raw_dir = os.path.join(out_dir, "raw")
        os.makedirs(gb.raw_dir, exist_ok=True)

    # Phase 2: enumerate all objects → nodes
    gb.scan_tables(app, db)
    gb.scan_relationships(db_path)
    gb.scan_queries(app, db)
    gb.scan_ui_objects(app)
    gb.finalize_data_names()

    # Build cross-module procedure index (must precede code heuristics)
    gb.scan_vba_projects(app)
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

    # Debug raw-export: keep raw files + rawHash/rawSize for every object,
    # including those skipped above when heuristics were disabled.
    if raw_export_mode == "debug":
        gb.export_raw_remaining(db_path, obj_list)

    # Snapshot last: the stamps must describe the design the graph was built from.
    try:
        gb.design_stamps = design_stamps(app, app.CurrentDb())
    except Exception as exc:
        gb.add_warning("DesignStampsFailed",
                       f"Could not record design timestamps: {exc}")

    # Phase 6: output
    return gb.build_output(abs_db, out_dir, field_mode, embed_viewer)
