import argparse
import json
import ssl
import sys
import traceback
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from extractor import extract_market_research
from renderer import render_market_research_pptx


ROOT = Path(__file__).resolve().parents[2]
OUTPUT_DIR = ROOT / "outputs" / "pptx_author_tool"
REFERENCE_PPT = ROOT / "reference_template.pptx"
SERVICE_NAME = "market-research-report-tool"


def openapi_spec(base_url):
    return {
        "openapi": "3.0.3",
        "info": {
            "title": "Market Research Report Tool",
            "version": "2.0.0",
            "description": (
                "Parse market-research Excel crosstabs into a traceable evidence package "
                "and render a structured slide specification into an editable PPTX report."
            ),
        },
        "servers": [{"url": base_url}],
        "paths": {
            "/extract-market-data": {
                "post": {
                    "operationId": "extract_market_research_data",
                    "summary": "Extract market research evidence",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "multipart/form-data": {
                                "schema": {
                                    "type": "object",
                                    "required": ["data_file"],
                                    "properties": {
                                        "data_file": {"type": "string", "format": "binary"},
                                        "business_focus": {"type": "string"},
                                        "max_evidence": {"type": "integer", "default": 1200},
                                    },
                                }
                            },
                            "application/json": {
                                "schema": {"$ref": "#/components/schemas/ExtractRequest"}
                            }
                        },
                    },
                    "responses": {
                        "200": {
                            "description": "ResearchEvidencePackage",
                            "content": {
                                "application/json": {
                                    "schema": {"$ref": "#/components/schemas/ExtractResponse"}
                                }
                            },
                        },
                        "400": {"description": "Invalid input file"},
                        "500": {"description": "Extraction failure"},
                    },
                }
            },
            "/render-pptx": {
                "post": {
                    "operationId": "render_market_research_pptx",
                    "summary": "Render market research PPTX",
                    "requestBody": {
                        "required": True,
                        "content": {
                            "application/json": {
                                "schema": {"$ref": "#/components/schemas/RenderRequest"}
                            }
                        },
                    },
                    "responses": {
                        "200": {
                            "description": "PPT render result",
                            "content": {
                                "application/json": {
                                    "schema": {"$ref": "#/components/schemas/RenderResponse"}
                                }
                            },
                        },
                        "400": {"description": "Invalid slide specification"},
                        "500": {"description": "Rendering failure"},
                    },
                }
            },
        },
        "components": {
            "schemas": {
                "DifyFile": {
                    "type": "object",
                    "description": "Dify file object with a reachable URL, path, or base64 content.",
                    "properties": {
                        "url": {"type": "string"},
                        "remote_url": {"type": "string"},
                        "path": {"type": "string"},
                        "content_base64": {"type": "string"},
                        "name": {"type": "string"},
                    },
                    "additionalProperties": True,
                },
                "ExtractRequest": {
                    "type": "object",
                    "required": ["data_file"],
                    "properties": {
                        "data_file": {
                            "oneOf": [
                                {"$ref": "#/components/schemas/DifyFile"},
                                {"type": "string"},
                            ]
                        },
                        "business_focus": {"type": "string"},
                        "max_evidence": {"type": "integer", "default": 1200},
                    },
                },
                "ExtractResponse": {
                    "type": "object",
                    "required": ["status", "schema_version", "research_evidence"],
                    "properties": {
                        "status": {"type": "string", "enum": ["success"]},
                        "schema_version": {"type": "string"},
                        "research_evidence": {"type": "object"},
                    },
                },
                "RenderRequest": {
                    "type": "object",
                    "required": ["slide_spec"],
                    "properties": {
                        "slide_spec": {"type": "object"},
                        "evidence_registry": {
                            "type": "array",
                            "items": {"type": "object"},
                        },
                        "output": {
                            "type": "object",
                            "properties": {
                                "format": {"type": "string", "enum": ["pptx"], "default": "pptx"},
                                "file_name": {"type": "string", "default": "market_research_report.pptx"},
                            },
                        },
                        "render_options": {
                            "type": "object",
                            "properties": {
                                "use_reference_dimensions": {"type": "boolean", "default": True},
                                "render_native_charts": {"type": "boolean", "default": True},
                                "include_source_notes": {"type": "boolean", "default": True},
                            },
                        },
                    },
                },
                "RenderResponse": {
                    "type": "object",
                    "required": ["status", "file_url", "file_name", "slide_count"],
                    "properties": {
                        "status": {"type": "string", "enum": ["success"]},
                        "schema_version": {"type": "string"},
                        "file_url": {"type": "string"},
                        "file_name": {"type": "string"},
                        "file_path": {"type": "string"},
                        "slide_count": {"type": "integer"},
                        "message": {"type": "string"},
                        "warnings": {"type": "array", "items": {"type": "string"}},
                        "visual_audit": {
                            "type": "object",
                            "description": "Structural and readability audit calculated from the rendered PPTX.",
                            "additionalProperties": True,
                        },
                    },
                },
            }
        },
    }


class ReportServer(ThreadingHTTPServer):
    daemon_threads = True
    scheme = "http"


class Handler(BaseHTTPRequestHandler):
    server_version = f"{SERVICE_NAME}/2.0"

    def _base_url(self):
        forwarded_proto = (self.headers.get("X-Forwarded-Proto") or "").split(",")[0].strip()
        forwarded_host = (self.headers.get("X-Forwarded-Host") or "").split(",")[0].strip()
        proto = forwarded_proto or self.server.scheme
        host = forwarded_host or self.headers.get("Host") or f"localhost:{self.server.server_port}"
        return f"{proto}://{host}"

    def _send_json(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _send_workflow_error(self, path, exc):
        message = str(exc)
        trace = traceback.format_exc(limit=5)
        if path == "/extract-market-data":
            self._send_json(
                200,
                {
                    "status": "error",
                    "schema_version": "research-evidence/1.0",
                    "message": message,
                    "trace": trace,
                    "research_evidence": {
                        "schema_version": "research-evidence/1.0",
                        "project": {},
                        "metric_definitions": [],
                        "evidence_registry": [],
                        "reference_report_outline": [],
                        "quality": {
                            "validation_errors": [message],
                            "normalized_evidence_count": 0,
                        },
                    },
                },
            )
            return True
        if path == "/render-pptx":
            self._send_json(
                200,
                {
                    "status": "error",
                    "schema_version": "render-response/1.0",
                    "message": message,
                    "trace": trace,
                    "file_url": "",
                    "file_name": "",
                    "slide_count": 0,
                    "visual_audit": {
                        "schema_version": "ppt-visual-audit/1.0",
                        "status": "failed",
                        "metrics": {
                            "code_like_text_hits": 0,
                            "out_of_bounds_shapes": 0,
                            "overlap_pairs": 0,
                        },
                        "slide_checks": [],
                    },
                },
            )
            return True
        return False

    def _read_json(self):
        raw = self._read_body()
        if not raw:
            return {}
        return json.loads(raw.decode("utf-8"))

    def _read_body(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            return b""
        if length > 50 * 1024 * 1024:
            raise ValueError("request body exceeds 50 MB")
        return self.rfile.read(length)

    def _read_multipart(self):
        content_type = self.headers.get("Content-Type", "")
        raw = self._read_body()
        if not raw:
            raise ValueError("multipart request body is empty")
        envelope = (
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
            + raw
        )
        message = BytesParser(policy=policy.default).parsebytes(envelope)
        if not message.is_multipart():
            raise ValueError("invalid multipart/form-data request")

        payload = {}
        for part in message.iter_parts():
            field_name = part.get_param("name", header="content-disposition")
            if not field_name:
                continue
            file_name = part.get_filename()
            content = part.get_payload(decode=True) or b""
            if file_name is not None:
                payload[field_name] = {
                    "content_bytes": content,
                    "name": file_name,
                    "content_type": part.get_content_type(),
                }
            else:
                charset = part.get_content_charset() or "utf-8"
                payload[field_name] = content.decode(charset)
        return payload

    def _read_payload(self):
        content_type = self.headers.get("Content-Type", "").lower()
        if content_type.startswith("multipart/form-data"):
            return self._read_multipart()
        if not content_type or content_type.startswith("application/json"):
            return self._read_json()
        raise ValueError(f"unsupported Content-Type: {content_type}")

    def _send_pptx(self, path):
        data = path.read_bytes()
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/openapi.json":
            self._send_json(200, openapi_spec(self._base_url()))
            return
        if path == "/health":
            self._send_json(
                200,
                {
                    "status": "ok",
                    "service": SERVICE_NAME,
                    "reference_ppt_available": REFERENCE_PPT.exists(),
                    "endpoints": ["/extract-market-data", "/render-pptx"],
                },
            )
            return
        if path.startswith("/files/"):
            file_name = Path(unquote(path)).name
            file_path = OUTPUT_DIR / file_name
            if not file_path.exists() or file_path.suffix.lower() != ".pptx":
                self._send_json(404, {"status": "error", "message": "file not found"})
                return
            self._send_pptx(file_path)
            return
        self._send_json(404, {"status": "error", "message": "not found"})

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            payload = self._read_payload()
            if path == "/extract-market-data":
                result = extract_market_research(payload, REFERENCE_PPT)
                self._send_json(200, result)
                return
            if path == "/render-pptx":
                result = render_market_research_pptx(
                    payload, OUTPUT_DIR, REFERENCE_PPT, self._base_url()
                )
                self._send_json(200, result)
                return
            self._send_json(404, {"status": "error", "message": "not found"})
        except (ValueError, json.JSONDecodeError) as exc:
            if not self._send_workflow_error(path, exc):
                self._send_json(400, {"status": "error", "message": str(exc)})
        except Exception as exc:
            if not self._send_workflow_error(path, exc):
                self._send_json(
                    500,
                    {
                        "status": "error",
                        "message": str(exc),
                        "trace": traceback.format_exc(limit=5),
                    },
                )

    def log_message(self, fmt, *args):
        sys.stderr.write("%s - %s\n" % (self.log_date_time_string(), fmt % args))


def main():
    parser = argparse.ArgumentParser(description="Dify market research report service")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8077)
    parser.add_argument("--certfile")
    parser.add_argument("--keyfile")
    args = parser.parse_args()

    server = ReportServer((args.host, args.port), Handler)
    if bool(args.certfile) != bool(args.keyfile):
        raise ValueError("--certfile and --keyfile must be provided together")
    if args.certfile and args.keyfile:
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(args.certfile, args.keyfile)
        server.socket = context.wrap_socket(server.socket, server_side=True)
        server.scheme = "https"

    print(f"{SERVICE_NAME} listening on {server.scheme}://{args.host}:{args.port}")
    print(f"Dify OpenAPI URL: {server.scheme}://{args.host}:{args.port}/openapi.json")
    server.serve_forever()


if __name__ == "__main__":
    main()
