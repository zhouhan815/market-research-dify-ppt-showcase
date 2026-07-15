import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import yaml


REQUIRED_NODE_FIELDS = {
    "start": {"title", "type", "variables"},
    "http-request": {
        "authorization",
        "body",
        "headers",
        "method",
        "params",
        "title",
        "type",
        "url",
    },
    "code": {"code", "code_language", "outputs", "title", "type", "variables"},
    "llm": {"context", "model", "prompt_template", "title", "type", "vision"},
    "if-else": {"cases", "title", "type"},
    "end": {"outputs", "title", "type"},
}

STANDARD_OUTPUTS = {
    "http-request": {"body", "files", "headers", "status_code"},
    "llm": {"reasoning_content", "text", "usage"},
}


def available_outputs(node):
    node_type = node["data"]["type"]
    if node_type == "start":
        return {item["variable"] for item in node["data"].get("variables", [])}
    if node_type == "code":
        return set(node["data"].get("outputs", {}))
    return STANDARD_OUTPUTS.get(node_type, set())


def main():
    root = Path(__file__).resolve().parents[2]
    workflow_path = (
        Path(sys.argv[1]).resolve()
        if len(sys.argv) > 1
        else root / "workflow" / "generic_market_research_ppt_workflow.yml"
    )
    document = yaml.safe_load(workflow_path.read_text(encoding="utf-8"))
    errors = []
    for dependency in document.get("dependencies", []):
        if dependency.get("type") != "marketplace":
            continue
        identifier = dependency.get("value", {}).get("marketplace_plugin_unique_identifier", "")
        if ":" not in identifier or "@" not in identifier:
            errors.append(f"incomplete marketplace plugin identifier: {identifier}")
    node_list = document["workflow"]["graph"]["nodes"]
    node_ids = [node["id"] for node in node_list]
    duplicate_node_ids = [node_id for node_id, count in Counter(node_ids).items() if count > 1]
    nodes = {node["id"]: node for node in node_list}
    edges = document["workflow"]["graph"]["edges"]

    for node_id in duplicate_node_ids:
        errors.append(f"duplicate node id: {node_id}")

    edge_ids = [edge["id"] for edge in edges]
    for edge_id, count in Counter(edge_ids).items():
        if count > 1:
            errors.append(f"duplicate edge id: {edge_id}")

    incoming = defaultdict(int)
    outgoing = defaultdict(int)

    for edge in edges:
        source = edge["source"]
        target = edge["target"]
        if source not in nodes:
            errors.append(f"missing source: {source}")
            continue
        if target not in nodes:
            errors.append(f"missing target: {target}")
            continue
        incoming[target] += 1
        outgoing[source] += 1
        if edge["data"]["sourceType"] != nodes[source]["data"]["type"]:
            errors.append(f"source type mismatch: {edge['id']}")
        if edge["data"]["targetType"] != nodes[target]["data"]["type"]:
            errors.append(f"target type mismatch: {edge['id']}")
        if nodes[source]["data"]["type"] == "if-else":
            valid_handles = {
                case["case_id"] for case in nodes[source]["data"].get("cases", [])
            } | {"false"}
            if edge.get("sourceHandle") not in valid_handles:
                errors.append(f"invalid if-else source handle: {edge['id']}")

    for node_id, node in nodes.items():
        data = node.get("data", {})
        node_type = data.get("type")
        if node_type not in REQUIRED_NODE_FIELDS:
            errors.append(f"unsupported node type: {node_id} ({node_type})")
            continue
        missing = sorted(REQUIRED_NODE_FIELDS[node_type] - set(data))
        if missing:
            errors.append(f"missing fields: {node_id}: {', '.join(missing)}")
        if node_type != "start" and incoming[node_id] == 0:
            errors.append(f"node has no incoming edge: {node_id}")
        if node_type != "end" and outgoing[node_id] == 0:
            errors.append(f"node has no outgoing edge: {node_id}")

        if node_type == "http-request":
            if not isinstance(data.get("params"), str):
                errors.append(f"http params must be a string: {node_id}")
            if not isinstance(data.get("headers"), str):
                errors.append(f"http headers must be a string: {node_id}")
            authorization = data.get("authorization") or {}
            if authorization.get("type") not in {"no-auth", "api-key"}:
                errors.append(f"invalid http authorization: {node_id}")
            body = data.get("body") or {}
            if body.get("type") not in {
                "none", "form-data", "x-www-form-urlencoded", "raw-text", "json", "binary"
            }:
                errors.append(f"invalid http body type: {node_id}")
            if body.get("type") == "form-data":
                fields = body.get("data")
                if not isinstance(fields, list) or not fields:
                    errors.append(f"form-data body must contain fields: {node_id}")
                    fields = []
                if "content-type" in data.get("headers", "").lower():
                    errors.append(f"form-data Content-Type boundary must be generated by Dify: {node_id}")
                for field in fields:
                    field_type = field.get("type")
                    if not field.get("key") or field_type not in {"text", "file"}:
                        errors.append(f"invalid form-data field: {node_id}: {field}")
                    if field_type == "file":
                        selector = field.get("file")
                        if not isinstance(selector, list) or len(selector) != 2:
                            errors.append(f"invalid form-data file selector: {node_id}: {selector}")
                        else:
                            source_id, output_name = selector
                            if source_id not in nodes or output_name not in available_outputs(nodes[source_id]):
                                errors.append(
                                    f"unknown form-data file selector: {node_id}: {source_id}.{output_name}"
                                )
                if node_id == "extract_market_data":
                    max_evidence_fields = [field for field in fields if field.get("key") == "max_evidence"]
                    if len(max_evidence_fields) != 1:
                        errors.append("extract_market_data must define one max_evidence field")
                    else:
                        try:
                            if int(max_evidence_fields[0].get("value", 0)) > 600:
                                errors.append("extract_market_data max_evidence must not exceed 600")
                        except (TypeError, ValueError):
                            errors.append("extract_market_data max_evidence must be an integer")

        if node_type == "code":
            declared = set(data.get("outputs", {}))
            if not declared:
                errors.append(f"code node has no outputs: {node_id}")
            for variable in data.get("variables", []):
                if not variable.get("variable") or not variable.get("value_selector"):
                    errors.append(f"invalid code input variable: {node_id}")

        if node_type == "llm":
            model = data.get("model") or {}
            for field in ("provider", "name", "mode"):
                if not model.get(field):
                    errors.append(f"missing llm model field: {node_id}.{field}")
            if not data.get("prompt_template"):
                errors.append(f"llm node has no prompt template: {node_id}")

        if node_type == "if-else":
            case_ids = [case.get("case_id") for case in data.get("cases", [])]
            if not case_ids or any(not case_id for case_id in case_ids):
                errors.append(f"if-else node has invalid cases: {node_id}")
            if len(case_ids) != len(set(case_ids)):
                errors.append(f"if-else node has duplicate case ids: {node_id}")

        if node_type == "end" and not data.get("outputs"):
            errors.append(f"end node has no outputs: {node_id}")

    env_outputs = {
        item.get("name")
        for item in document.get("workflow", {}).get("environment_variables", [])
        if item.get("name")
    }

    workflow_text = workflow_path.read_text(encoding="utf-8")
    references = re.findall(r"\{\{#([A-Za-z0-9_]+)\.([A-Za-z0-9_]+)#\}\}", workflow_text)
    for node_id, variable in references:
        if node_id == "env":
            if variable not in env_outputs:
                errors.append(f"unknown environment variable: env.{variable}")
            continue
        if node_id not in nodes:
            errors.append(f"unknown variable node: {node_id}.{variable}")
            continue
        if variable not in available_outputs(nodes[node_id]):
            errors.append(f"unknown node output: {node_id}.{variable}")

    for node_id, node in nodes.items():
        data = node["data"]
        selectors = []
        if data["type"] == "code":
            selectors.extend(item.get("value_selector", []) for item in data.get("variables", []))
        if data["type"] == "if-else":
            selectors.extend(
                condition.get("variable_selector", [])
                for case in data.get("cases", [])
                for condition in case.get("conditions", [])
            )
        if data["type"] == "end":
            selectors.extend(item.get("value_selector", []) for item in data.get("outputs", []))
        for selector in selectors:
            if len(selector) != 2:
                errors.append(f"invalid value selector: {node_id}: {selector}")
                continue
            source_id, output_name = selector
            if source_id not in nodes:
                errors.append(f"unknown selector node: {node_id}: {source_id}.{output_name}")
            elif output_name not in available_outputs(nodes[source_id]):
                errors.append(f"unknown selector output: {node_id}: {source_id}.{output_name}")

    code_results = {}
    for node_id, node in nodes.items():
        if node["data"]["type"] != "code":
            continue
        try:
            compile(node["data"]["code"], f"<{node_id}>", "exec")
            code_results[node_id] = "ok"
        except Exception as exc:
            code_results[node_id] = str(exc)
            errors.append(f"code compile failed: {node_id}: {exc}")

    result = {
        "app": document["app"]["name"],
        "start_variables": [item["variable"] for item in nodes["start"]["data"]["variables"]],
        "nodes": len(node_list),
        "edges": len(edges),
        "code_nodes": code_results,
        "errors": errors,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(1 if errors else 0)


if __name__ == "__main__":
    main()
