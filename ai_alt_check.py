"""
AI review of alt text, using a local vision model through Ollama.

Rule-based checks can tell that an image has alt text, but not whether the
alt text describes the image. This module asks a local model to compare each
image with its alt text and flag the ones a person should look at. It never
passes or fails anything on its own: every flag is a prompt for human review.

The model runs locally (https://ollama.com), so no images leave the machine.
Results are cached by image, alt text and model, so a rescan only asks about
images that changed.
"""

import base64
import hashlib
import io
import json
import os
import urllib.error
import urllib.request

from PIL import Image

DEFAULT_MODEL = "gemma3"
DEFAULT_HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434")

# Vector formats PowerPoint uses that Pillow can't rasterize; these are skipped.
UNSUPPORTED_TYPES = {"image/x-emf", "image/x-wmf", "image/emf", "image/wmf", "image/svg+xml"}

PROMPT = """You are helping an accessibility specialist review alternative text (alt text) for images in a slide presentation.

Slide title: {title}
Other text on the slide: {context}
The image's alt text: "{alt}"

Look at the image. Decide whether the alt text gives a screen reader user what they need from this image in this context.

Answer "needs_review" if the alt text:
- describes something different from what the image shows;
- misses information the image exists to convey, such as words, numbers, or the trend in a chart;
- is a placeholder, a file name, or a generic label like "image" or "graph".

Answer "adequate" if it conveys the image's content or purpose. Short alt text is fine if it does that. Don't flag style, length or wording preferences.

Reply with JSON only: {{"verdict": "adequate" or "needs_review", "reason": "one short sentence"}}"""


class AltTextReviewer:
    def __init__(self, model=DEFAULT_MODEL, host=DEFAULT_HOST, cache_path=None, timeout=120):
        self.model = model
        self.host = host.rstrip("/")
        self.timeout = timeout
        self.cache_path = cache_path
        self.cache = {}
        self.stats = {"reviewed": 0, "flagged": 0, "cached": 0, "skipped": 0, "errors": 0}
        if cache_path and os.path.exists(cache_path):
            try:
                with open(cache_path, encoding="utf-8") as f:
                    self.cache = json.load(f)
            except (OSError, ValueError):
                self.cache = {}

    def check_available(self):
        """Return None if Ollama is reachable and has the model, else an error message."""
        try:
            with urllib.request.urlopen(f"{self.host}/api/tags", timeout=10) as resp:
                tags = json.load(resp)
        except (urllib.error.URLError, OSError, ValueError) as e:
            return f"Can't reach Ollama at {self.host} ({e}). Is `ollama serve` running?"
        names = {m.get("name", "") for m in tags.get("models", [])}
        if not any(n == self.model or n.split(":")[0] == self.model for n in names):
            return f'Model "{self.model}" isn\'t installed. Run: ollama pull {self.model}'
        return None

    def review(self, image_blob, content_type, alt_text, title="", context=""):
        """Return None if the alt text looks adequate (or couldn't be checked),
        otherwise a short reason to show a reviewer."""
        if content_type in UNSUPPORTED_TYPES:
            self.stats["skipped"] += 1
            return None

        key = hashlib.sha256(image_blob + b"\0" + alt_text.encode() + b"\0" + self.model.encode()).hexdigest()
        if key in self.cache:
            self.stats["cached"] += 1
            result = self.cache[key]
        else:
            try:
                png = _to_png(image_blob)
            except Exception:
                self.stats["skipped"] += 1
                return None
            result = self._ask(png, alt_text, title, context)
            if result is None:
                self.stats["errors"] += 1
                return None
            self.cache[key] = result
            self.stats["reviewed"] += 1

        if result.get("verdict") == "needs_review":
            self.stats["flagged"] += 1
            return result.get("reason") or "The alt text may not describe this image."
        return None

    def _ask(self, png, alt_text, title, context):
        prompt = PROMPT.format(
            title=title or "(none)",
            context=(context or "(none)")[:600],
            alt=alt_text.replace('"', "'")[:500],
        )
        body = json.dumps({
            "model": self.model,
            "prompt": prompt,
            "images": [base64.b64encode(png).decode()],
            "format": "json",
            "stream": False,
            "options": {"temperature": 0},
        }).encode()
        req = urllib.request.Request(
            f"{self.host}/api/generate", data=body, headers={"Content-Type": "application/json"}
        )
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                answer = json.loads(json.load(resp).get("response", ""))
        except (urllib.error.URLError, OSError, ValueError):
            return None
        verdict = str(answer.get("verdict", "")).strip().lower()
        if verdict not in ("adequate", "needs_review"):
            return None
        return {"verdict": verdict, "reason": str(answer.get("reason", "")).strip()[:300]}

    def save(self):
        if not self.cache_path:
            return
        try:
            with open(self.cache_path, "w", encoding="utf-8") as f:
                json.dump(self.cache, f)
        except OSError as e:
            print(f"Warning: could not save AI review cache: {e}")


def _to_png(blob, max_side=1024):
    """Downscale large images so requests stay small; the model doesn't need full resolution."""
    img = Image.open(io.BytesIO(blob))
    img.thumbnail((max_side, max_side))
    if img.mode not in ("RGB", "RGBA", "L"):
        img = img.convert("RGBA")
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()
