"""
graph_query.py — Query an Access dependency graph without re-scanning.

Loads a previously-generated graph.json and supports targeted queries:
  neighbors  — direct connections to/from a node (depth 1-3)
  impact     — transitive downstream dependents (what breaks if this changes)
  path       — shortest path between two nodes
  orphans    — nodes with zero incoming edges
  summary    — high-level stats, top edge kinds, high-degree nodes
"""

from __future__ import annotations

import json
import os
from collections import defaultdict, deque
from typing import Any


# ---------------------------------------------------------------------------
# Graph loader
# ---------------------------------------------------------------------------

class _Graph:
    """Lightweight in-memory graph built from graph.json."""

    __slots__ = ("nodes", "edges", "out_adj", "in_adj", "meta",
                 "_id_lookup", "_label_lookup")

    def __init__(self, data: dict) -> None:
        self.meta: dict = data.get("meta", {})
        self.nodes: dict[str, dict] = {}
        self.edges: list[dict] = data.get("edges", [])
        self.out_adj: dict[str, list[dict]] = defaultdict(list)
        self.in_adj: dict[str, list[dict]] = defaultdict(list)
        self._id_lookup: dict[str, dict] = {}
        self._label_lookup: dict[str, list[dict]] = defaultdict(list)

        for n in data.get("nodes", []):
            nid = n["id"]
            self.nodes[nid] = n
            self._id_lookup[nid.lower()] = n
            self._label_lookup[n["label"].lower()].append(n)

        for e in self.edges:
            self.out_adj[e["from"]].append(e)
            self.in_adj[e["to"]].append(e)

    # ── node resolution ─────────────────────────────────────────────────

    def resolve(self, name: str) -> list[dict]:
        """Resolve a user-supplied name to node(s).

        Priority: exact id > group:name > label match.
        """
        if not name:
            return []

        # 1. Exact id
        low = name.strip().lower()
        if low in self._id_lookup:
            return [self._id_lookup[low]]

        # 2. Try group:name for all groups
        _GROUPS = ("table", "query", "form", "report", "macro",
                   "module", "field", "sql")
        for g in _GROUPS:
            candidate = f"{g}:{name}".lower()
            if candidate in self._id_lookup:
                return [self._id_lookup[candidate]]

        # 3. Label match (case-insensitive)
        hits = self._label_lookup.get(low, [])
        if hits:
            return list(hits)

        return []


def _load_graph(graph_path: str | None, db_path: str | None) -> _Graph:
    """Load graph.json, auto-locating it next to the database if needed."""
    path = graph_path
    if not path and db_path:
        candidate = os.path.join(
            os.path.dirname(os.path.abspath(db_path)),
            "access-graph-out", "graph.json",
        )
        if os.path.isfile(candidate):
            path = candidate

    if not path or not os.path.isfile(path):
        raise FileNotFoundError(
            f"graph.json not found. "
            f"Run access_graph first to generate the dependency graph. "
            f"Searched: {path or '(no path provided)'}"
        )

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return _Graph(data)


# ---------------------------------------------------------------------------
# Compact edge/node formatters
# ---------------------------------------------------------------------------

def _fmt_node(n: dict) -> dict:
    out: dict[str, Any] = {"id": n["id"], "label": n["label"], "group": n["group"]}
    meta = n.get("meta", {})
    if "sqlPreview" in meta:
        out["sqlPreview"] = meta["sqlPreview"]
    if "fieldCount" in meta:
        out["fieldCount"] = meta["fieldCount"]
    return out


def _fmt_edge(e: dict) -> dict:
    return {
        "from": e["from"],
        "to": e["to"],
        "kind": e["kind"],
        "label": e["label"],
    }


# ---------------------------------------------------------------------------
# Action implementations
# ---------------------------------------------------------------------------

_MAX_RESULTS = 200


def _action_neighbors(
    g: _Graph, node_id: str, depth: int, direction: str,
    skip_fields: bool,
) -> dict:
    """BFS neighbors up to `depth` hops."""
    depth = max(1, min(depth, 3))
    visited: set[str] = {node_id}
    result_in: list[dict] = []
    result_out: list[dict] = []

    # BFS outgoing
    if direction in ("out", "both"):
        frontier: deque[tuple[str, int]] = deque([(node_id, 0)])
        while frontier:
            cur, d = frontier.popleft()
            if d >= depth:
                continue
            for e in g.out_adj.get(cur, []):
                target = e["to"]
                if skip_fields and e["kind"] == "field-owner":
                    continue
                result_out.append({**_fmt_edge(e), "depth": d + 1})
                if target not in visited:
                    visited.add(target)
                    frontier.append((target, d + 1))

    # BFS incoming
    if direction in ("in", "both"):
        frontier = deque([(node_id, 0)])
        visited_in: set[str] = {node_id}
        while frontier:
            cur, d = frontier.popleft()
            if d >= depth:
                continue
            for e in g.in_adj.get(cur, []):
                source = e["from"]
                if skip_fields and e["kind"] == "field-owner":
                    continue
                result_in.append({**_fmt_edge(e), "depth": d + 1})
                if source not in visited_in:
                    visited_in.add(source)
                    frontier.append((source, d + 1))

    truncated_in = len(result_in) > _MAX_RESULTS
    truncated_out = len(result_out) > _MAX_RESULTS

    return {
        "action": "neighbors",
        "node": _fmt_node(g.nodes[node_id]),
        "depth": depth,
        "direction": direction,
        "incoming": result_in[:_MAX_RESULTS],
        "outgoing": result_out[:_MAX_RESULTS],
        "total_incoming": len(result_in),
        "total_outgoing": len(result_out),
        "truncated": truncated_in or truncated_out,
    }


def _action_impact(g: _Graph, node_id: str, skip_fields: bool) -> dict:
    """Transitive downstream walk — everything that depends on this node."""
    visited: set[str] = {node_id}
    frontier: deque[str] = deque([node_id])
    affected: list[dict] = []
    edges_used: list[dict] = []

    while frontier:
        cur = frontier.popleft()
        for e in g.out_adj.get(cur, []):
            if skip_fields and e["kind"] == "field-owner":
                continue
            target = e["to"]
            edges_used.append(_fmt_edge(e))
            if target not in visited:
                visited.add(target)
                affected.append(_fmt_node(g.nodes[target]))
                frontier.append(target)

        if len(affected) >= _MAX_RESULTS:
            break

    truncated = len(affected) >= _MAX_RESULTS

    # Group by type for readability
    by_group: dict[str, list[str]] = defaultdict(list)
    for a in affected:
        by_group[a["group"]].append(a["label"])

    return {
        "action": "impact",
        "node": _fmt_node(g.nodes[node_id]),
        "affected_count": len(affected),
        "affected_by_group": dict(by_group),
        "affected": affected[:_MAX_RESULTS],
        "edges": edges_used[:_MAX_RESULTS],
        "truncated": truncated,
    }


def _action_path(g: _Graph, source_id: str, target_id: str) -> dict:
    """BFS shortest path (undirected — either edge direction counts)."""
    if source_id == target_id:
        return {
            "action": "path",
            "source": _fmt_node(g.nodes[source_id]),
            "target": _fmt_node(g.nodes[target_id]),
            "found": True,
            "path": [_fmt_node(g.nodes[source_id])],
            "edges": [],
            "length": 0,
        }

    # Build undirected adjacency for BFS
    adj: dict[str, list[tuple[str, dict]]] = defaultdict(list)
    for e in g.edges:
        adj[e["from"]].append((e["to"], e))
        adj[e["to"]].append((e["from"], e))

    visited: set[str] = {source_id}
    # parent map: child_id -> (parent_id, edge)
    parent: dict[str, tuple[str, dict]] = {}
    frontier: deque[str] = deque([source_id])
    found = False

    while frontier:
        cur = frontier.popleft()
        for neighbor, edge in adj[cur]:
            if neighbor not in visited:
                visited.add(neighbor)
                parent[neighbor] = (cur, edge)
                if neighbor == target_id:
                    found = True
                    break
                frontier.append(neighbor)
        if found:
            break

    if not found:
        return {
            "action": "path",
            "source": _fmt_node(g.nodes[source_id]),
            "target": _fmt_node(g.nodes[target_id]),
            "found": False,
            "path": [],
            "edges": [],
            "length": -1,
        }

    # Reconstruct path
    path_nodes: list[dict] = []
    path_edges: list[dict] = []
    cur = target_id
    while cur != source_id:
        path_nodes.append(_fmt_node(g.nodes[cur]))
        p, e = parent[cur]
        path_edges.append(_fmt_edge(e))
        cur = p
    path_nodes.append(_fmt_node(g.nodes[source_id]))
    path_nodes.reverse()
    path_edges.reverse()

    return {
        "action": "path",
        "source": _fmt_node(g.nodes[source_id]),
        "target": _fmt_node(g.nodes[target_id]),
        "found": True,
        "path": path_nodes,
        "edges": path_edges,
        "length": len(path_edges),
    }


def _action_orphans(g: _Graph, skip_fields: bool) -> dict:
    """Nodes with zero incoming edges (potential dead objects)."""
    orphans: list[dict] = []
    for nid, node in g.nodes.items():
        if skip_fields and node["group"] == "field":
            continue
        incoming = g.in_adj.get(nid, [])
        if skip_fields:
            incoming = [e for e in incoming if e["kind"] != "field-owner"]
        if not incoming:
            orphans.append(_fmt_node(node))

    # Group by type
    by_group: dict[str, list[str]] = defaultdict(list)
    for o in orphans:
        by_group[o["group"]].append(o["label"])

    return {
        "action": "orphans",
        "count": len(orphans),
        "by_group": dict(by_group),
        "orphans": orphans[:_MAX_RESULTS],
        "truncated": len(orphans) > _MAX_RESULTS,
    }


def _action_summary(g: _Graph, group: str | None) -> dict:
    """High-level stats: counts, top edge kinds, high-degree nodes."""
    # Node counts by group
    node_counts: dict[str, int] = defaultdict(int)
    for n in g.nodes.values():
        node_counts[n["group"]] += 1

    # Edge counts by kind
    edge_counts: dict[str, int] = defaultdict(int)
    for e in g.edges:
        edge_counts[e["kind"]] += 1

    # Degree analysis (combined in+out, excluding field-owner)
    degree: dict[str, int] = defaultdict(int)
    for e in g.edges:
        if e["kind"] == "field-owner":
            continue
        degree[e["from"]] += 1
        degree[e["to"]] += 1

    # Top 15 by degree
    top = sorted(degree.items(), key=lambda x: x[1], reverse=True)[:15]
    top_nodes = []
    for nid, deg in top:
        node = g.nodes.get(nid)
        if node:
            if group and node["group"] != group:
                continue
            top_nodes.append({**_fmt_node(node), "degree": deg})

    # Filter nodes/edges if group specified
    filtered_edge_counts = edge_counts
    if group:
        filtered_edge_counts = defaultdict(int)
        for e in g.edges:
            src = g.nodes.get(e["from"], {})
            tgt = g.nodes.get(e["to"], {})
            if src.get("group") == group or tgt.get("group") == group:
                filtered_edge_counts[e["kind"]] += 1

    sorted_edges = sorted(filtered_edge_counts.items(),
                          key=lambda x: x[1], reverse=True)

    return {
        "action": "summary",
        "group_filter": group,
        "node_counts": dict(node_counts),
        "total_nodes": len(g.nodes),
        "total_edges": len(g.edges),
        "edge_kinds": sorted_edges,
        "top_connected_nodes": top_nodes[:15],
    }


# ---------------------------------------------------------------------------
# Entry point (called from dispatcher — no COM needed)
# ---------------------------------------------------------------------------

def ac_graph_query(
    action: str,
    graph_path: str | None = None,
    db_path: str | None = None,
    node: str | None = None,
    source: str | None = None,
    target: str | None = None,
    depth: int = 1,
    direction: str = "both",
    group: str | None = None,
    skip_fields: bool = True,
) -> dict:
    """Query a previously-generated Access dependency graph.

    Actions:
        neighbors — direct connections to/from a node
        impact    — transitive downstream dependents
        path      — shortest path between two nodes
        orphans   — nodes with zero incoming edges
        summary   — high-level stats and top-degree nodes
    """
    g = _load_graph(graph_path, db_path)

    # --- resolve nodes for actions that need them ---
    def _must_resolve(name: str | None, param: str) -> str:
        if not name:
            raise ValueError(f"'{param}' is required for action '{action}'")
        hits = g.resolve(name)
        if not hits:
            raise ValueError(
                f"Node '{name}' not found in graph. "
                f"Try the full id (e.g. 'table:Customers') or check spelling."
            )
        if len(hits) > 1:
            options = [f"{h['id']} ({h['group']})" for h in hits[:10]]
            raise ValueError(
                f"'{name}' is ambiguous — matches: {', '.join(options)}. "
                f"Use the full id to disambiguate."
            )
        return hits[0]["id"]

    if action == "neighbors":
        nid = _must_resolve(node, "node")
        return _action_neighbors(g, nid, depth, direction, skip_fields)

    elif action == "impact":
        nid = _must_resolve(node, "node")
        return _action_impact(g, nid, skip_fields)

    elif action == "path":
        sid = _must_resolve(source, "source")
        tid = _must_resolve(target, "target")
        return _action_path(g, sid, tid)

    elif action == "orphans":
        return _action_orphans(g, skip_fields)

    elif action == "summary":
        return _action_summary(g, group)

    else:
        raise ValueError(
            f"Unknown action '{action}'. "
            f"Valid actions: neighbors, impact, path, orphans, summary"
        )
