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
    GraphBuilder, _embedded_macro_blocks, _extract_record_source,
    _qualified_field_refs, _strip_vba_comments,
)
from mcp_access import graph_query  # noqa: E402
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


def test_build_output_records_design_stamps_and_warning_counts(tmp_path):
    gb = _builder_with("module:m")
    gb.design_stamps = {"module:m": "2026-09-01 10:00:00"}
    gb._warn_missing("module:m", "form", "frmGone", "VBA OpenForm")
    r = gb.build_output(str(tmp_path / "x.accdb"), str(tmp_path / "out"),
                        "referenced", embed_viewer=False)
    assert r["warnings_by_code"] == {"MissingReference": 1}
    meta = json.load(open(r["graph_path"], encoding="utf-8"))["meta"]
    assert meta["designStamps"] == {"module:m": "2026-09-01 10:00:00"}


# ---------------------------------------------------------------------------
# broken action / freshness
# ---------------------------------------------------------------------------

def _graph_with_warnings(tmp_path, stamps=None) -> str:
    gb = _builder_with("module:modNav", "form:frmMain")
    gb._warn_missing("module:modNav", "form", "frmOld", "VBA OpenForm")
    gb._warn_missing("form:frmMain", "form", "frmOld", "SourceObject of sf")
    gb.add_warning("ExportFailed", "boom", {"owner": "rptX", "group": "report"})
    path = os.path.join(str(tmp_path), "graph.json")
    meta = {"warnings": gb.warnings, "generatedAt": "2026-09-29T00:00:00+00:00"}
    if stamps is not None:
        meta["designStamps"] = stamps
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


STAMPS = {"table:T": "2026-09-01", "form:F": "2026-09-02", "query:Q": "2026-09-03"}


def test_freshness_fresh_when_design_unchanged(tmp_path, monkeypatch):
    monkeypatch.setattr(graph_query, "_live_design_stamps", lambda db: dict(STAMPS))
    r = ac_graph_query("summary", graph_path=_graph_with_warnings(tmp_path, STAMPS))
    assert r["graph"]["stale"] is False
    assert r["graph"]["generatedAt"] == "2026-09-29T00:00:00+00:00"


def test_freshness_lists_changed_added_removed(tmp_path, monkeypatch):
    live = {"table:T": "2026-09-01", "form:F": "2026-09-29", "module:New": "x"}
    monkeypatch.setattr(graph_query, "_live_design_stamps", lambda db: live)
    r = ac_graph_query("summary", graph_path=_graph_with_warnings(tmp_path, STAMPS))
    fresh = r["graph"]
    assert fresh["stale"] is True
    assert fresh["changed"] == ["form:F"]
    assert fresh["added"] == ["module:New"]
    assert fresh["removed"] == ["query:Q"]
    assert "access_graph" in fresh["note"]


def test_freshness_unknown_when_database_not_open(tmp_path):
    r = ac_graph_query("summary", graph_path=_graph_with_warnings(tmp_path, STAMPS))
    assert r["graph"]["stale"] is None
    assert "not open" in r["graph"]["note"]


def test_freshness_unknown_for_old_graphs(sample):
    r = ac_graph_query("summary", graph_path=sample)
    assert r["graph"]["stale"] is None
    assert "predates" in r["graph"]["note"]


# ---------------------------------------------------------------------------
# event properties, expressions, embedded macros
# ---------------------------------------------------------------------------

FORM_EXPORT = '''Version =21
Begin Form
    RecordSource ="Customers"
    OnLoad ="=InitForm()"
    Begin
        Begin Label
            BackStyle =0
        End
    End
    Begin Section
        Begin
            Begin CommandButton
                OnClick ="mcrMenu.OpenOrders"
                Name ="cmdMenu"
            End
            Begin CommandButton
                Name ="cmdSave"
                OnClick ="[Event Procedure]"
            End
            Begin CommandButton
                OnClick ="[Embedded Macro]"
                OnClickEmMacro = Begin
                    Version =196611
                    Begin
                        Action ="OpenForm"
                        Argument ="frmOrders"
                        Argument ="0"
                    End
                    Begin
                        Action ="OpenReport"
                        Argument ="rptGone"
                    End
                End
                Name ="cmdEmb"
            End
            Begin TextBox
                Name ="txtTotal"
                ControlSource ="=CalcTotal([ID]) & [Forms]![frmOrders]![txtX]"
            End
            Begin CommandButton
                Name ="cmdGone"
                OnDblClick ="mcrMissing"
            End
        End
    End
End
'''


def _edges(gb: GraphBuilder, kind: str) -> list[dict]:
    return [e for e in gb.edges if e["kind"] == kind]


def test_embedded_macro_blocks_resolve_owner_named_after_the_block():
    blocks = _embedded_macro_blocks(FORM_EXPORT)
    assert [(b["control"], b["property"]) for b in blocks] == [("cmdEmb", "OnClick")]
    assert any('Action ="OpenForm"' in ln for ln in blocks[0]["lines"])


def test_event_properties_expressions_and_embedded_macros():
    gb = _builder_with("form:frmMain", "form:frmOrders", "macro:mcrMenu",
                       "module:modUtil")
    gb.index_module_procs(
        "module:modUtil",
        "Public Function InitForm()\nEnd Function\nFunction CalcTotal(x)\nEnd Function\n")
    gb._analyze_properties("form:frmMain", FORM_EXPORT, None)

    calls = {(e["kind"], e["meta"]["procedure"], e["meta"].get("controlName"))
             for e in gb.edges if "procedure" in e["meta"]}
    assert calls == {("event-call", "InitForm", None),
                     ("expression-call", "CalcTotal", "txtTotal")}

    [macro_edge] = _edges(gb, "event-macro")
    assert macro_edge["to"] == "macro:mcrMenu"
    assert macro_edge["meta"]["controlName"] == "cmdMenu"

    [emb] = _edges(gb, "macro-openform")
    assert emb["to"] == "form:frmOrders"
    assert emb["meta"]["embedded"] == "cmdEmb.OnClick"

    [ref] = _edges(gb, "form-reference")
    assert ref["to"] == "form:frmOrders"
    assert ref["meta"]["targetControl"] == "txtX"

    missing = {(w["meta"]["targetGroup"], w["meta"]["target"]) for w in _missing(gb)}
    assert missing == {("report", "rptGone"), ("macro", "mcrMissing")}


def test_macro_runmacro_and_runcode():
    gb = _builder_with("macro:mcrMain", "macro:mcrOther", "module:modUtil")
    gb.index_module_procs("module:modUtil", "Function DoStuff()\nEnd Function\n")
    lines = ['Action ="RunMacro"', 'Argument ="mcrOther.Sub1"',
             'Action ="RunCode"', 'Argument ="DoStuff()"',
             'Action ="RunMacro"', 'Argument ="mcrMain.Self"']
    gb._analyze_macro_lines("macro:mcrMain", lines, None)
    assert [e["to"] for e in _edges(gb, "macro-runmacro")] == ["macro:mcrOther"]
    [runcode] = _edges(gb, "macro-runcode")
    assert runcode["to"] == "module:modUtil"
    assert runcode["meta"]["procedure"] == "DoStuff"
    assert _missing(gb) == []


# ---------------------------------------------------------------------------
# VBA calls and references
# ---------------------------------------------------------------------------

def _calls_builder() -> GraphBuilder:
    gb = _builder_with("module:modA", "module:modB", "module:clsThing")
    gb.index_module_procs(
        "module:modB",
        "Public Sub DoWork(x)\nEnd Sub\nPublic Function GetVal()\nEnd Function\n"
        "Sub Helper()\nEnd Sub\n")
    gb.index_module_procs("module:clsThing", "Public Sub Init()\nEnd Sub\n",
                          is_class=True)
    gb._compile_proc_call_re()
    return gb


def test_sub_calls_without_parentheses_are_detected():
    gb = _calls_builder()
    code = (
        "Sub Main()\n"
        "    DoWork 1, 2\n"
        "    If x Then Helper\n"
        '    MsgBox "GetVal() is text, not a call"\n'
        "    Init\n"
        "End Sub\n"
        "Private Sub Other()\n"
        "    y = 1: GetVal = 5\n"
        "End Sub\n"
    )
    gb._analyze_code_heuristics("module:modA", "module", "modA", code, None)
    assert {e["meta"]["procedure"] for e in _edges(gb, "vba-call")} == {
        "DoWork", "Helper"}
    assert all(e["to"] == "module:modB" for e in _edges(gb, "vba-call"))


def test_local_procedure_shadows_public_one():
    gb = _calls_builder()
    code = "Private Sub DoWork()\nEnd Sub\nSub Main()\n    DoWork\n    Call DoWork\nEnd Sub\n"
    gb._analyze_code_heuristics("module:modA", "module", "modA", code, None)
    assert _edges(gb, "vba-call") == []


def test_vba_execute_openrecordset_and_forms_refs():
    gb = _builder_with("module:m", "query:qryAppend", "form:frmA")
    gb.finalize_data_names()
    code = (
        'CurrentDb.Execute "qryAppend", dbFailOnError\n'
        'Set rs = db.OpenRecordset("SELECT * FROM qryAppend")\n'
        'Forms("frmA").Requery\n'
        'x = Forms!frmGone!txtA\n'
    )
    gb._analyze_code_heuristics("module:m", "module", "m", code, None)
    assert [e["to"] for e in _edges(gb, "vba-data-ref")
            if e["label"] == "Execute"] == ["query:qryAppend"]
    assert any(e["label"] == "OpenRecordset" for e in _edges(gb, "vba-runsql"))
    assert [e["to"] for e in _edges(gb, "form-reference")] == ["form:frmA"]
    assert [w["meta"]["target"] for w in _missing(gb)] == ["frmGone"]


def test_forms_reference_in_sql():
    gb = _builder_with("query:q", "form:frmA")
    gb._link_object_refs(
        "query:q",
        "SELECT * FROM T WHERE ID=[Forms]![frmA]![cboID] AND X=Forms!frmGone.txt",
        {"via": "SQL"})
    [ref] = _edges(gb, "form-reference")
    assert ref["to"] == "form:frmA" and ref["meta"]["targetControl"] == "cboID"
    assert [w["meta"]["target"] for w in _missing(gb)] == ["frmGone"]


# ---------------------------------------------------------------------------
# field lineage
# ---------------------------------------------------------------------------

def test_qualified_field_refs_with_aliases_and_strings():
    tf = {"Customers": {"ID": "Long", "Phone": "Text"},
          "Order Details": {"Qty": "Integer"}}
    sql = ('SELECT c.Phone, [Order Details].[Qty], Customers.ID '
           'FROM Customers AS c INNER JOIN [Order Details] '
           'ON c.ID = [Order Details].OrderID '
           'WHERE Customers.Nope = "x.Phone"')
    assert _qualified_field_refs(sql, tf) == [
        ("Customers", "Phone"), ("Order Details", "Qty"), ("Customers", "ID")]


def _lineage_builder() -> GraphBuilder:
    gb = _builder_with("table:Customers", "query:qryCust", "form:frmCust")
    gb._known_table_fields["Customers"] = {"ID": "Long", "Phone": "Text"}
    gb.set_query_fields("qryCust", [
        ("Phone", "Customers", "Phone", "Text"),
        ("Calc", "", "", "Double"),
        ("Customers.ID", "Customers", "ID", "Long"),
    ])
    gb.finalize_data_names()
    return gb


def test_query_bound_control_links_through_lineage_to_table_field(tmp_path):
    gb = _lineage_builder()
    rs = gb._resolve_record_source("form:frmCust", "form", "frmCust", "qryCust", None)
    gb._handle_control_source("form:frmCust", "phone", "txtPhone", "TextBox", rs, None)
    gb._handle_control_source("form:frmCust", "Customers.ID", "txtID", "TextBox", rs, None)

    [lineage] = [e for e in _edges(gb, "field-lineage")
                 if e["from"] == "field:query:qryCust:Phone"]
    assert lineage["to"] == "field:table:Customers:Phone"
    assert gb.nodes["field:query:qryCust:Phone"]["meta"]["verified"] is True
    assert gb.warnings == []

    path = _save(gb, tmp_path)
    r = ac_graph_query("impact", graph_path=path, node="field:table:Customers:Phone")
    assert [(a["id"], a["depth"]) for a in r["affected"]] == [("form:frmCust", 2)]
    r = ac_graph_query("impact", graph_path=path, node="table:Customers")
    assert "form:frmCust" in {a["id"] for a in r["affected"]}


def test_control_bound_to_missing_field_is_warned():
    gb = _lineage_builder()
    rs = gb._resolve_record_source("form:frmCust", "form", "frmCust", "qryCust", None)
    gb._handle_control_source("form:frmCust", "Fax", "txtFax", "TextBox", rs, None)
    [w] = gb.warnings
    assert w["code"] == "MissingField"
    assert (w["meta"]["target"], w["meta"]["field"]) == ("qryCust", "Fax")


def test_single_table_sql_record_source_binds_table_fields():
    gb = _builder_with("table:Customers", "form:F")
    gb._known_table_fields["Customers"] = {"ID": "Long", "Phone": "Text"}
    gb.finalize_data_names()
    sql = "SELECT Customers.Phone, Customers.ID * 2 AS Dbl FROM Customers"
    rs = gb._resolve_record_source("form:F", "form", "F", sql, None)
    assert rs["table"] == "Customers"
    gb._handle_control_source("form:F", "Phone", "txtPhone", "TextBox", rs, None)
    gb._handle_control_source("form:F", "Dbl", "txtDbl", "TextBox", rs, None)
    assert [e["to"] for e in _edges(gb, "controlsource")] == [
        "field:table:Customers:Phone"]
    assert {e["to"] for e in _edges(gb, "sql-field")} == {
        "field:table:Customers:Phone", "field:table:Customers:ID"}
    assert gb.warnings == []


def test_unverified_reference_never_downgrades_a_verified_field():
    gb = _builder_with("table:T")
    gb._ensure_field_node("table:T", "table", "T", "X", True, "Text")
    gb._ensure_field_node("table:T", "table", "T", "X", False, None)
    meta = gb.nodes["field:table:T:X"]["meta"]
    assert meta["verified"] is True and meta["dataType"] == "Text"
