"""Configurazione: impostazioni da variabili d'ambiente (prefisso AD_) e profili aziendali JSON."""
from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _load_dotenv(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, _, v = line.partition("=")
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


@dataclass
class Settings:
    host: str = "127.0.0.1"
    port: int = 8000
    profiles_dir: Path = ROOT / "profiles"
    rounding: str = "half_up"          # half_up | half_even | down | up
    decimals: int = 2
    margin_threshold: Decimal = Decimal("20")   # % sotto la quale il margine è evidenziato
    vat_default: Decimal = Decimal("22")        # solo come valore proposto, sempre modificabile
    max_upload_mb: int = 25
    llm_provider: str = "none"         # none | anthropic
    llm_api_key: str = ""
    llm_model: str = "claude-sonnet-5-5"
    llm_base_url: str = "https://api.anthropic.com"

    @property
    def llm_enabled(self) -> bool:
        return self.llm_provider != "none" and bool(self.llm_api_key)

    @classmethod
    def from_env(cls) -> "Settings":
        _load_dotenv(ROOT / ".env")
        e = os.environ.get
        return cls(
            host=e("AD_HOST", "127.0.0.1"),
            port=int(e("AD_PORT", "8000")),
            profiles_dir=Path(e("AD_PROFILES_DIR", str(ROOT / "profiles"))),
            rounding=e("AD_ROUNDING", "half_up"),
            decimals=int(e("AD_DECIMALS", "2")),
            margin_threshold=Decimal(e("AD_MARGIN_THRESHOLD", "20")),
            vat_default=Decimal(e("AD_VAT_DEFAULT", "22")),
            max_upload_mb=int(e("AD_MAX_UPLOAD_MB", "25")),
            llm_provider=e("AD_LLM_PROVIDER", "none").lower(),
            llm_api_key=e("AD_LLM_API_KEY", ""),
            llm_model=e("AD_LLM_MODEL", "claude-sonnet-5-5"),
            llm_base_url=e("AD_LLM_BASE_URL", "https://api.anthropic.com"),
        )


@dataclass
class Profile:
    id: str
    name: str
    company_name: str = "AD Consulting S.p.A."
    fonts: dict = field(default_factory=lambda: {"body": "Calibri", "heading": "Calibri"})
    sizes: dict = field(default_factory=lambda: {"body": 11, "h1": 16, "h2": 13, "h3": 12, "title": 26, "table": 9.5})
    palette: dict = field(default_factory=lambda: {
        "primary": "1F3A5F", "secondary": "2E7D9A", "text": "222222", "accent": "C8102E",
        "table_header_text": "FFFFFF", "table_band": "EEF3F8"})
    margins_cm: dict = field(default_factory=lambda: {"top": 2.5, "bottom": 2.2, "left": 2.3, "right": 2.3})
    logo: str | None = None
    logo_path: Path | None = None
    cover: bool = False
    header_text: str = ""
    footer_text: str = ""
    required_sections: list = field(default_factory=list)   # [{"title":..., "aliases":[...]}]
    required_fields: list = field(default_factory=lambda: [
        "cliente", "numero_offerta", "data_offerta", "oggetto", "perimetro", "prezzi",
        "durata", "rinnovo", "validita", "pagamento"])
    table_headers_expected: list = field(default_factory=list)
    typos: dict = field(default_factory=dict)
    known_entities: list = field(default_factory=list)
    margin_threshold_pct: Decimal | None = None
    raw: dict = field(default_factory=dict)

    @classmethod
    def from_file(cls, path: Path) -> "Profile":
        d = json.loads(path.read_text(encoding="utf-8"))
        known = {k: v for k, v in d.items() if k in cls.__dataclass_fields__ and k not in ("raw", "logo_path")}
        p = cls(**known)
        if d.get("margin_threshold_pct") is not None:
            p.margin_threshold_pct = Decimal(str(d["margin_threshold_pct"]))
        if p.logo:
            lp = (path.parent / p.logo)
            p.logo_path = lp if lp.exists() else None
        p.raw = d
        return p


def load_profiles(settings: Settings) -> dict[str, Profile]:
    out: dict[str, Profile] = {}
    if settings.profiles_dir.exists():
        for f in sorted(settings.profiles_dir.glob("*.json")):
            try:
                p = Profile.from_file(f)
                out[p.id] = p
            except Exception as exc:  # profilo malformato: non blocca l'avvio
                print(f"[profili] {f.name} ignorato: {exc}")
    if not out:
        out["default"] = Profile(id="default", name="Profilo predefinito")
    return out
