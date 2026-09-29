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

from mcp_access.graph import GraphBuilder, _extract_record_source  # noqa: E402
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
