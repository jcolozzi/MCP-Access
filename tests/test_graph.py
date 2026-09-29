"""
Pure-Python tests for the dependency graph (graph.py / graph_query.py).

No COM / no Access needed: graphs are built with GraphBuilder primitives and
queried through ac_graph_query against a temp graph.json.

Run with:
    python -m pytest tests/test_graph.py
"""

import json
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access.graph import (  # noqa: E402
    GraphBuilder, _extract_record_source, _strip_vba_comments,
)
from mcp_access.graph_query import ac_graph_query  # noqa: E402


def _save(gb: GraphBuilder, tmp_path) -> str:
    path = os.path.join(str(tmp_path), "graph.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"meta": {}, "nodes": list(gb.nodes.values()),
                   "edges": gb.edges}, f)
    return path


@pytest.fixture
def sample(tmp_path):
    """Customers <- qryCust <- frmCust <- frmMain (OpenForm); frmMain binds Customers.Phone."""
    gb = GraphBuilder()
    gb.add_node("table:Customers", "Customers", "table", is_data=True)
    gb.add_node("table:Orders", "Orders", "table", is_data=True)
    gb.add_node("query:qryCust", "qryCust", "query", is_data=True)
    gb.add_node("form:frmCust", "frmCust", "form")
    gb.add_node("form:frmMain", "frmMain", "form")
    gb.add_edge("query:qryCust", "table:Customers", "Customers",
                "query-sql-reference")
    gb.add_edge("form:frmCust", "query:qryCust", "RecordSource", "recordsource")
    gb.add_edge("form:frmMain", "form:frmCust", "OpenForm", "vba-openform",
                meta={"name": "frmCust"})
    fid = gb._ensure_field_node("table:Customers", "table", "Customers",
                                "Phone", True, "Text")
    gb.add_edge("form:frmMain", fid, "ControlSource", "controlsource",
                meta={"controlName": "txtPhone"})
    return _save(gb, tmp_path)


# ---------------------------------------------------------------------------
# impact: dependents, not dependencies
# ---------------------------------------------------------------------------

def test_impact_returns_dependents_of_a_table(sample):
    r = ac_graph_query("impact", graph_path=sample, node="Customers")
    labels = {a["label"] for a in r["affected"]}
    assert labels == {"qryCust", "frmCust", "frmMain"}


def test_impact_does_not_return_dependencies(sample):
    r = ac_graph_query("impact", graph_path=sample, node="frmCust")
    labels = {a["label"] for a in r["affected"]}
    assert labels == {"frmMain"}
    assert "Customers" not in labels and "qryCust" not in labels


def test_impact_reaches_controls_bound_to_a_tables_fields(tmp_path):
    gb = GraphBuilder()
    gb.add_node("table:T", "T", "table", is_data=True)
    gb.add_node("form:F", "F", "form")
    fid = gb._ensure_field_node("table:T", "table", "T", "X", True, "Text")
    gb.add_edge("form:F", fid, "ControlSource", "controlsource",
                meta={"controlName": "txtX"})
    r = ac_graph_query("impact", graph_path=_save(gb, tmp_path), node="table:T")
    assert [a["id"] for a in r["affected"]] == ["form:F"]
    assert r["affected"][0]["depth"] == 1
    assert r["edges"][0]["meta"] == {"controlName": "txtX"}


def test_impact_skip_fields_false_lists_the_fields(sample):
    r = ac_graph_query("impact", graph_path=sample, node="Customers",
                       skip_fields=False)
    assert "field" in r["affected_by_group"]


def test_impact_on_a_field_does_not_climb_to_its_owner(sample):
    r = ac_graph_query("impact", graph_path=sample,
                       node="field:table:Customers:Phone")
    assert {a["label"] for a in r["affected"]} == {"frmMain"}


def test_impact_reports_depth(sample):
    r = ac_graph_query("impact", graph_path=sample, node="Customers")
    depth = {a["label"]: a["depth"] for a in r["affected"]}
    assert depth["qryCust"] == 1
    assert depth["frmCust"] == 2
    # frmMain binds a Customers field directly, which beats the OpenForm chain.
    assert depth["frmMain"] == 1


def test_impact_of_unreferenced_table_is_empty(sample):
    r = ac_graph_query("impact", graph_path=sample, node="Orders")
    assert r["affected_count"] == 0


# ---------------------------------------------------------------------------
# output carries what an agent needs to act
# ---------------------------------------------------------------------------

def test_edges_carry_meta(sample):
    r = ac_graph_query("neighbors", graph_path=sample, node="frmMain",
                       direction="out")
    by_kind = {e["kind"]: e for e in r["outgoing"]}
    assert by_kind["controlsource"]["meta"]["controlName"] == "txtPhone"
    assert by_kind["vba-openform"]["meta"] == {"name": "frmCust"}


def test_sql_nodes_expose_preview_and_origin(tmp_path):
    gb = GraphBuilder()
    gb.add_node("form:F", "F", "form")
    sid = gb._ensure_sql_node("SELECT * FROM T", "form:F:RecordSource")
    gb.add_edge("form:F", sid, "RecordSource", "recordsource-sql")
    r = ac_graph_query("neighbors", graph_path=_save(gb, tmp_path),
                       node=sid, direction="in")
    assert r["node"]["preview"] == "SELECT * FROM T"
    assert r["node"]["origin"] == "form:F:RecordSource"


# ---------------------------------------------------------------------------
# resolve / summary
# ---------------------------------------------------------------------------

def test_same_name_in_two_groups_is_ambiguous(tmp_path):
    gb = GraphBuilder()
    gb.add_node("table:Customers", "Customers", "table", is_data=True)
    gb.add_node("form:Customers", "Customers", "form")
    path = _save(gb, tmp_path)
    with pytest.raises(ValueError, match="ambiguous"):
        ac_graph_query("impact", graph_path=path, node="Customers")
    r = ac_graph_query("impact", graph_path=path, node="form:Customers")
    assert r["node"]["id"] == "form:Customers"


def test_summary_group_filter_is_applied_before_top_n(tmp_path):
    gb = GraphBuilder()
    gb.add_node("form:F", "F", "form")
    for i in range(20):
        gb.add_node(f"table:T{i}", f"T{i}", "table", is_data=True)
        gb.add_node(f"query:Q{i}", f"Q{i}", "query", is_data=True)
        for j in range(5):
            gb.add_edge(f"query:Q{i}", f"table:T{i}", str(j), f"k{j}")
    gb.add_edge("form:F", "table:T0", "RecordSource", "recordsource")
    r = ac_graph_query("summary", graph_path=_save(gb, tmp_path), group="form")
    assert [n["id"] for n in r["top_connected_nodes"]] == ["form:F"]


# ---------------------------------------------------------------------------
# RecordSource parsing
# ---------------------------------------------------------------------------

def test_record_source_joins_wrapped_lines():
    text = (
        "Version =21\nBegin Form\n"
        '    RecordSource ="SELECT Customers.ID FROM Customers WHERE Cu"\n'
        '        "stomers.Active = True"\n'
        '    Caption ="x"\nBegin Section\n'
    )
    assert _extract_record_source(text) == (
        "SELECT Customers.ID FROM Customers WHERE Customers.Active = True"
    )


def test_record_source_keeps_its_own_closing_quote():
    text = ('Begin Form\n    RecordSource ="SELECT * FROM T WHERE a=\\"x\\""\n'
            "Begin Section\n")
    assert _extract_record_source(text) == 'SELECT * FROM T WHERE a=\\"x\\"'


def test_record_source_plain_name_and_absent():
    assert _extract_record_source(
        'Begin Form\n    RecordSource ="Customers"\nBegin Section\n'
    ) == "Customers"
    assert _extract_record_source("Begin Form\nBegin Section\n") is None


def test_record_source_after_first_section_is_ignored():
    text = 'Begin Form\nBegin Section\n    RecordSource ="Other"\n'
    assert _extract_record_source(text) is None


# ---------------------------------------------------------------------------
# viewer embedding
# ---------------------------------------------------------------------------

def test_viewer_embed_cannot_close_the_script_tag(tmp_path):
    gb = GraphBuilder()
    gb.add_node("query:q", "q", "query",
                meta={"sqlPreview": 'SELECT "</script><script>alert(1)</script>"'})
    viewer = gb._write_viewer(str(tmp_path), {
        "meta": {}, "nodes": list(gb.nodes.values()), "edges": []})
    assert viewer is not None
    html = open(viewer, encoding="utf-8").read()
    start = html.index("var EMBEDDED_GRAPH = ")
    embedded = html[start:html.index("</script>", start)]
    assert "alert(1)" in embedded  # payload stayed inside the data
    data = json.loads(embedded[len("var EMBEDDED_GRAPH = "):].rstrip().rstrip(";"))
    assert data["nodes"][0]["meta"]["sqlPreview"].startswith('SELECT "</script>')


# ---------------------------------------------------------------------------
# VBA comment stripping
# ---------------------------------------------------------------------------

def test_strip_vba_comments():
    code = (
        "DoCmd.OpenForm \"frmA\" ' open A\n"
        "' DoCmd.OpenForm \"frmOld\"\n"
        "Rem DoCmd.OpenForm \"frmRem\"\n"
        "MsgBox \"it's \"\"quoted\"\" here\" ' tail\n"
        "Remark = 1"
    )
    out = _strip_vba_comments(code).split("\n")
    assert len(out) == 5
    assert out[0].rstrip() == 'DoCmd.OpenForm "frmA"'
    assert out[1] == "" and out[2] == ""
    assert out[3].rstrip() == 'MsgBox "it\'s ""quoted"" here"'
    assert out[4] == "Remark = 1"


# ---------------------------------------------------------------------------
# missing references
# ---------------------------------------------------------------------------

def _builder_with(*ids: str) -> GraphBuilder:
    gb = GraphBuilder()
    for nid in ids:
        group, name = nid.split(":", 1)
        gb.add_node(nid, name, group, is_data=group in ("table", "query"))
    return gb


def _missing(gb: GraphBuilder) -> list[dict]:
    return [w for w in gb.warnings if w["code"] == "MissingReference"]


def test_vba_reference_to_missing_form_is_warned_and_existing_is_linked():
    gb = _builder_with("module:modNav", "form:frmA")
    code = 'DoCmd.OpenForm "frmA"\nDoCmd.OpenForm "frmGone"\nDoCmd.OpenForm "frmGone"\n'
    gb._analyze_code_heuristics("module:modNav", "module", "modNav", code, None)
    assert [(e["to"], e["kind"]) for e in gb.edges] == [("form:frmA", "vba-openform")]
    missing = _missing(gb)
    assert len(missing) == 1  # deduplicated
    assert missing[0]["meta"]["target"] == "frmGone"
    assert missing[0]["meta"]["targetGroup"] == "form"
    assert missing[0]["meta"]["ownerId"] == "module:modNav"


def test_commented_out_code_creates_no_edge_or_warning():
    gb = _builder_with("module:m", "form:frmA")
    code = "' DoCmd.OpenForm \"frmA\"\nRem DoCmd.OpenForm \"frmGone\"\n"
    gb._analyze_code_heuristics("module:m", "module", "m", code, None)
    assert gb.edges == [] and _missing(gb) == []


def test_runmacro_submacro_resolves_to_macro_group():
    gb = _builder_with("module:m", "macro:mcrMenu")
    gb._analyze_code_heuristics(
        "module:m", "module", "m", 'DoCmd.RunMacro "mcrMenu.OpenOrders"\n', None)
    assert [e["to"] for e in gb.edges] == ["macro:mcrMenu"]
    assert _missing(gb) == []


def test_source_object_missing_is_warned_and_table_prefix_resolves():
    gb = _builder_with("form:frmMain", "table:tblLog")
    gb._handle_source_object("form:frmMain", "Table.tblLog", "sfLog", "SubForm", {})
    gb._handle_source_object("form:frmMain", "Form.sfGone", "sfX", "SubForm", {})
    assert [e["to"] for e in gb.edges] == ["table:tblLog"]
    missing = _missing(gb)
    assert len(missing) == 1
    assert missing[0]["meta"]["target"] == "Form.sfGone"
    assert "sfX" in missing[0]["meta"]["via"]


def test_expression_targets_are_not_reported_missing():
    gb = _builder_with("macro:m")
    gb._warn_missing("macro:m", "form", "=[Forms]![x]", "macro OpenForm")
    assert _missing(gb) == []


def test_build_output_records_mtime_and_warning_counts(tmp_path):
    db = tmp_path / "x.accdb"
    db.write_bytes(b"")
    gb = _builder_with("module:m")
    gb._warn_missing("module:m", "form", "frmGone", "VBA OpenForm")
    r = gb.build_output(str(db), str(tmp_path / "out"), "referenced",
                        embed_viewer=False)
    assert r["warnings_by_code"] == {"MissingReference": 1}
    meta = json.load(open(r["graph_path"], encoding="utf-8"))["meta"]
    assert meta["databaseMtime"] == pytest.approx(os.path.getmtime(db))


# ---------------------------------------------------------------------------
# broken action / freshness
# ---------------------------------------------------------------------------

def _graph_with_warnings(tmp_path, db_mtime=None, db=None) -> str:
    gb = _builder_with("module:modNav", "form:frmMain")
    gb._warn_missing("module:modNav", "form", "frmOld", "VBA OpenForm")
    gb._warn_missing("form:frmMain", "form", "frmOld", "SourceObject of sf")
    gb.add_warning("ExportFailed", "boom", {"owner": "rptX", "group": "report"})
    path = os.path.join(str(tmp_path), "graph.json")
    meta = {"warnings": gb.warnings, "generatedAt": "2026-09-29T00:00:00+00:00"}
    if db is not None:
        meta["database"] = db
    if db_mtime is not None:
        meta["databaseMtime"] = db_mtime
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "nodes": list(gb.nodes.values()),
                   "edges": []}, f)
    return path


def test_broken_lists_all_warnings(tmp_path):
    r = ac_graph_query("broken", graph_path=_graph_with_warnings(tmp_path))
    assert r["count"] == 3
    assert r["by_code"] == {"MissingReference": 2, "ExportFailed": 1}


def test_broken_filters_by_deleted_target_name(tmp_path):
    r = ac_graph_query("broken", graph_path=_graph_with_warnings(tmp_path),
                       node="FRMOLD")
    assert r["count"] == 2
    assert {w["meta"]["owner"] for w in r["warnings"]} == {"modNav", "frmMain"}


def test_broken_filters_by_owner(tmp_path):
    r = ac_graph_query("broken", graph_path=_graph_with_warnings(tmp_path),
                       node="rptX")
    assert [w["code"] for w in r["warnings"]] == ["ExportFailed"]


def test_freshness_fresh_and_stale(tmp_path):
    db = tmp_path / "x.accdb"
    db.write_bytes(b"")
    built = os.path.getmtime(db)
    path = _graph_with_warnings(tmp_path, db_mtime=built, db=str(db))
    r = ac_graph_query("summary", graph_path=path)
    assert r["graph"]["stale"] is False
    assert r["graph"]["generatedAt"] == "2026-09-29T00:00:00+00:00"

    os.utime(db, (built + 60, built + 60))
    r = ac_graph_query("summary", graph_path=path)
    assert r["graph"]["stale"] is True
    assert "access_graph" in r["graph"]["note"]


def test_freshness_unknown_for_old_graphs(sample):
    r = ac_graph_query("summary", graph_path=sample)
    assert r["graph"]["stale"] is None
