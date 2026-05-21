import json
from typing import Any, Generator
from urllib.parse import urljoin
from urllib.request import Request, urlopen

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage


class ForecastFormulaReadbackTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        base_url = self.runtime.credentials.get("base_url") or "http://localhost:4321"
        address = str(tool_parameters.get("address") or "B3").upper()
        sheet_name = str(tool_parameters.get("sheet_name") or "Inputs")
        value = tool_parameters.get("value", 0.4)

        try:
            proof = call_bilig_forecast(base_url=base_url, sheet_name=sheet_name, address=address, value=value)
            yield self.create_json_message(json=compact_proof(proof))
        except Exception as error:
            yield self.create_text_message(text=f"Bilig WorkPaper forecast readback failed: {error}")


def call_bilig_forecast(
    *,
    base_url: str,
    address: str,
    value: Any,
    sheet_name: str = "Inputs",
    timeout: int = 30,
) -> dict[str, Any]:
    if not base_url.startswith(("http://", "https://")):
        raise ValueError("base_url must start with http:// or https://")

    endpoint = urljoin(base_url.rstrip("/") + "/", "api/workpaper/n8n/forecast")
    payload = json.dumps({"sheetName": sheet_name, "address": address, "value": value}).encode("utf-8")
    request = Request(
        endpoint,
        data=payload,
        headers={"content-type": "application/json", "accept": "application/json"},
        method="POST",
    )

    with urlopen(request, timeout=timeout) as response:
        body = response.read().decode("utf-8")

    parsed = json.loads(body)
    if not isinstance(parsed, dict):
        raise ValueError(f"Expected JSON object response, received {type(parsed).__name__}")
    if parsed.get("verified") is not True:
        raise ValueError(f"Unverified WorkPaper response: {parsed}")
    return parsed


def compact_proof(proof: dict[str, Any]) -> dict[str, Any]:
    checks = proof.get("checks") if isinstance(proof.get("checks"), dict) else {}
    before = proof.get("before") if isinstance(proof.get("before"), dict) else {}
    after = proof.get("after") if isinstance(proof.get("after"), dict) else {}

    return {
        "verified": proof.get("verified") is True,
        "editedCell": proof.get("editedCell"),
        "before": {
            "expectedArr": before.get("expectedArr"),
            "targetGap": before.get("targetGap"),
        },
        "after": {
            "expectedArr": after.get("expectedArr"),
            "targetGap": after.get("targetGap"),
        },
        "checks": {
            "formulasPersisted": checks.get("formulasPersisted") is True,
            "restoredMatchesAfter": checks.get("restoredMatchesAfter") is True,
            "computedOutputChanged": checks.get("computedOutputChanged") is True,
        },
        "source": "Bilig WorkPaper",
        "github": "https://github.com/proompteng/bilig",
    }
