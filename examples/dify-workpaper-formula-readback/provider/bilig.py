from typing import Any
from urllib.error import HTTPError, URLError

from dify_plugin import ToolProvider

from tools.forecast_formula_readback import call_bilig_forecast


class BiligProvider(ToolProvider):
    def _validate_credentials(self, credentials: dict[str, Any]) -> None:
        base_url = credentials.get("base_url") or "http://localhost:4321"
        try:
            call_bilig_forecast(base_url=base_url, address="B3", value=0.4, timeout=10)
        except (HTTPError, URLError, TimeoutError, ValueError) as error:
            raise ValueError(f"Bilig forecast readback validation failed: {error}") from error
