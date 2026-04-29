from __future__ import annotations

import os
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Optional

import ollama

try:
    import win32com.client  # type: ignore
except Exception:
    win32com = None  # type: ignore

try:
    import screen_brightness_control as sbc  # type: ignore
except Exception:
    sbc = None  # type: ignore

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume  # type: ignore
    from ctypes import POINTER, cast
except Exception:
    AudioUtilities = None  # type: ignore
    IAudioEndpointVolume = None  # type: ignore
    POINTER = None  # type: ignore
    cast = None  # type: ignore


@dataclass
class JarvisResult:
    text: str
    action: Optional[Dict[str, Any]] = None


class JarvisAssistant:
    def __init__(self, model: str = "phi", enable_tts: bool = False) -> None:
        self.model = model
        self.enable_tts = enable_tts
        self._speaker = None
        if self.enable_tts and win32com is not None:
            try:
                self._speaker = win32com.client.Dispatch("SAPI.SpVoice")
            except Exception:
                self._speaker = None

    def speak(self, text: str) -> None:
        if self._speaker is not None:
            self._speaker.Speak(text, 1)

    def save_response_to_file(self, response: str) -> None:
        with open("jarvis_ai_answers.txt", "a", encoding="utf-8") as file:
            time_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            file.write(f"[{time_now}] {response}\n")

    def _set_volume(self, percent: int) -> bool:
        if AudioUtilities is None or IAudioEndpointVolume is None or cast is None or POINTER is None:
            return False
        devices = AudioUtilities.GetSpeakers()
        interface = devices.EndpointVolume
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(percent / 100, None)
        return True

    def _set_brightness(self, percent: int) -> bool:
        if sbc is None:
            return False
        sbc.set_brightness(percent)
        return True

    def _handle_local_command(self, query: str) -> Optional[JarvisResult]:
        lowered = query.lower().strip()

        volume_match = re.search(r"volume (\d+)", lowered)
        if volume_match:
            percent = max(0, min(100, int(volume_match.group(1))))
            changed = self._set_volume(percent)
            return JarvisResult(
                text=f"Volume set to {percent}%." if changed else "Volume control is not available on this system."
            )

        brightness_match = re.search(r"brightness (\d+)", lowered)
        if brightness_match:
            percent = max(0, min(100, int(brightness_match.group(1))))
            changed = self._set_brightness(percent)
            return JarvisResult(
                text=f"Brightness set to {percent}%." if changed else "Brightness control is not available on this system."
            )

        if lowered.startswith("open "):
            site_name = lowered.replace("open ", "", 1).strip()
            if site_name:
                url = f"https://www.{site_name}.com"
                return JarvisResult(text=f"Opening {site_name}.", action={"type": "open_url", "url": url})

        if "the time" in lowered or lowered == "time":
            return JarvisResult(text=f"The time is {datetime.now().strftime('%H:%M:%S')}.")

        if "date" in lowered:
            return JarvisResult(text=f"Today's date is {datetime.now().strftime('%A, %d %B %Y')}.")

        return None

    def ask_model(self, prompt: str) -> str:
        response = ollama.chat(model=self.model, messages=[{"role": "user", "content": prompt}])
        output = response["message"]["content"]
        self.save_response_to_file(output)
        return output

    def process_text(self, query: str) -> JarvisResult:
        query = (query or "").strip()
        if not query:
            return JarvisResult(text="Please say or type something.")

        command_result = self._handle_local_command(query)
        if command_result is not None:
            self.speak(command_result.text)
            return command_result

        try:
            output = self.ask_model(query)
            self.speak(output)
            return JarvisResult(text=output)
        except Exception as exc:
            return JarvisResult(text=f"Model error: {exc}")
