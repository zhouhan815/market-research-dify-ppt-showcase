import json
import sys
import threading
import uuid
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from server import Handler, ReportServer


def request_json(url, payload=None):
    data = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = Request(url, data=data, headers={"Content-Type": "application/json"})
    try:
        with urlopen(request, timeout=120) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} from {url}: {body}") from exc


def request_multipart(url, fields, file_field, file_path):
    boundary = "----DifyWorkflowTest" + uuid.uuid4().hex
    chunks = []
    for key, value in fields.items():
        chunks.extend([
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode(),
            str(value).encode("utf-8"),
            b"\r\n",
        ])
    chunks.extend([
        f"--{boundary}\r\n".encode(),
        (
            f'Content-Disposition: form-data; name="{file_field}"; '
            f'filename="{file_path.name}"\r\n'
        ).encode("utf-8"),
        b"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n",
        file_path.read_bytes(),
        b"\r\n",
        f"--{boundary}--\r\n".encode(),
    ])
    request = Request(
        url,
        data=b"".join(chunks),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
    )
    try:
        with urlopen(request, timeout=120) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} from {url}: {body}") from exc


def main():
    server = ReportServer(("127.0.0.1", 0), Handler)
    port = server.server_address[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    base_url = f"http://127.0.0.1:{port}"

    try:
        health = request_json(base_url + "/health")
        extracted = request_multipart(
            base_url + "/extract-market-data",
            fields={
                "business_focus": "重点关注购买意愿、价值感和目标人群",
                "max_evidence": 200,
            },
            file_field="data_file",
            file_path=ROOT / "sample_input" / "synthetic_concept_test_data.xlsx",
        )
        rendered = request_json(
            base_url + "/render-pptx",
            {
                "slide_spec": {
                    "schema_version": "market-research-slides/1.0",
                    "slides": [
                        {
                            "slide_id": "S-001",
                            "page_no": 1,
                            "section": "封面",
                            "title": "DemoHealthNutritionLineMulti-VitaminMulti-Mineral",
                            "layout": "cover",
                            "body": {
                                "key_message": "概念测试调研报告",
                                "supporting_points": [],
                                "implication": "",
                                "recommended_action": "",
                            },
                            "visual": {"type": "none"},
                            "source_notes": [],
                        },
                        {
                            "slide_id": "S-002",
                            "page_no": 2,
                            "section": "主要发现",
                            "title": "两款概念购买意愿接近，Concept B方向性更高",
                            "layout": "kpi_comparison",
                            "body": {
                                "key_message": "",
                                "supporting_points": [],
                                "implication": "继续验证价格价值感",
                                "recommended_action": "强化Concept B的价值沟通",
                            },
                            "visual": {
                                "type": "kpi_comparison",
                                "chart_type": "clustered_bar",
                                "categories": ["肯定会买", "可能会买"],
                                "series": [
                                    {"name": "Concept A", "values": [46, 44]},
                                    {"name": "Concept B", "values": [51, 39]},
                                ],
                            },
                            "source_notes": ["Q4-1，Header1，Base=200"],
                        },
                    ],
                },
                "output": {"file_name": "service_http_demo_test.pptx"},
            },
        )
        result = {
            "health": health,
            "extract": {
                "status": extracted["status"],
                "schema_version": extracted["schema_version"],
                "evidence_count": len(extracted["research_evidence"]["evidence_registry"]),
                "reference_outline_pages": len(extracted["research_evidence"]["reference_report_outline"]),
            },
            "render": rendered,
        }
        print(json.dumps(result, ensure_ascii=False, indent=2))
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=5)


if __name__ == "__main__":
    main()
