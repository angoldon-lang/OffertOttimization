import io
import re
import subprocess
import sys
import zipfile
from decimal import Decimal
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
DEMO = ROOT / "demo"


@pytest.fixture(scope="session", autouse=True)
def demo_files():
    if not (DEMO / "offerta_triennale_conforme.docx").exists():
        subprocess.run([sys.executable, str(ROOT / "scripts" / "make_demo.py")], check=True)


@pytest.fixture()
def service():
    from adreview.config import Settings, load_profiles
    st = Settings.from_env()
    st.llm_provider = "none"
    return __import__("adreview.service", fromlist=["Service"]).Service(st, load_profiles(st))


def demo(name: str) -> bytes:
    return (DEMO / name).read_bytes()


def docx_text(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        xml = " ".join(z.read(n).decode("utf8", "ignore") for n in z.namelist() if n.startswith("word/") and n.endswith(".xml"))
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", xml))


@pytest.fixture()
def mk(service):
    def _mk(offer, template=None, quotes=(), internal=None, **kw):
        return service.create_session(offer, demo(offer), "demo", (template, demo(template)) if template else None,
                                      [(q, demo(q)) for q in quotes], internal, **kw)
    return _mk
