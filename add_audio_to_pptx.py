"""
Add ElevenLabs TTS audio narration to a PowerPoint file.

Requirements:
    pip install python-pptx requests lxml

Usage:
    python add_audio_to_pptx.py

The script reads speaker notes from each slide, generates MP3 audio via
ElevenLabs, saves each audio file locally, then embeds them into the PPTX.
"""

import os
import time
import requests
from lxml import etree
from pptx import Presentation
from pptx.opc.part import Part
from pptx.opc.packuri import PackURI

# ── Configuration ────────────────────────────────────────────────────────────
ELEVENLABS_API_KEY = "3edf18bfe491b6d56480eb4b801091d3b735823f28858f8d9080194f6abed971"
VOICE_ID = "JBFqnCBsd6RMkjVDRZzb"  # "George" — change to any ElevenLabs voice ID
MODEL_ID = "eleven_multilingual_v2"

INPUT_PPTX  = "training_week_6_with_narration (2).pptx"
OUTPUT_PPTX = "training_week_6_with_audio.pptx"
AUDIO_DIR   = "slide_audio"   # folder to save MP3 files
# ─────────────────────────────────────────────────────────────────────────────

AUDIO_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
PPTNS = "http://schemas.openxmlformats.org/presentationml/2006/main"
ANS   = "http://schemas.openxmlformats.org/drawingml/2006/main"
RNS   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def get_slide_notes(slide) -> str:
    if slide.has_notes_slide:
        return slide.notes_slide.notes_text_frame.text.strip()
    return ""


def generate_audio(text: str) -> bytes:
    """Call ElevenLabs TTS API and return MP3 bytes."""
    url = f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE_ID}"
    headers = {
        "xi-api-key": ELEVENLABS_API_KEY,
        "Content-Type": "application/json",
        "Accept": "audio/mpeg",
    }
    payload = {
        "text": text,
        "model_id": MODEL_ID,
        "voice_settings": {"stability": 0.5, "similarity_boost": 0.75},
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    return resp.content


def embed_audio(prs: Presentation, slide, audio_bytes: bytes, slide_index: int):
    """Embed MP3 bytes into a slide and add auto-play timing."""
    # Create a proper Part object for the audio
    partname = PackURI(f"/ppt/media/audio{slide_index + 1}.mp3")
    audio_part = Part(partname, "audio/mpeg", audio_bytes, slide.part.package)
    rId = slide.part.relate_to(audio_part, AUDIO_REL_TYPE)

    shape_id = 100 + slide_index

    # Hidden picture shape that holds the audio
    pic_xml = (
        f'<p:pic xmlns:p="{PPTNS}" xmlns:a="{ANS}" xmlns:r="{RNS}">'
        f'<p:nvPicPr>'
        f'<p:cNvPr id="{shape_id}" name="Audio{slide_index+1}"/>'
        f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
        f'<p:nvPr><p:audioFile r:link="{rId}"/></p:nvPr>'
        f'</p:nvPicPr>'
        f'<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</p:spPr>'
        f'</p:pic>'
    )
    slide.shapes._spTree.append(etree.fromstring(pic_xml))

    # Auto-play timing
    timing_xml = (
        f'<p:timing xmlns:p="{PPTNS}" xmlns:a="{ANS}">'
        f'<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="whenNotActive" nodeType="tmRoot">'
        f'<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par>'
        f'<p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
        f'<p:childTnLst><p:par><p:cTn id="4" fill="hold">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst><p:audio>'
        f'<p:cMediaNode vol="80000" mute="0" showWhenStopped="0">'
        f'<p:cTn id="5" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        f'<p:tgtEl><p:spTgt spid="{shape_id}"/></p:tgtEl>'
        f'</p:cMediaNode></p:audio></p:childTnLst>'
        f'</p:cTn></p:par></p:childTnLst></p:cTn>'
        f'</p:par></p:childTnLst></p:cTn>'
        f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst/></p:timing>'
    )
    slide_el = slide._element
    existing = slide_el.find(f"{{{PPTNS}}}timing")
    if existing is not None:
        slide_el.remove(existing)
    slide_el.append(etree.fromstring(timing_xml))


def main():
    os.makedirs(AUDIO_DIR, exist_ok=True)
    print(f"Loading: {INPUT_PPTX}")
    prs = Presentation(INPUT_PPTX)
    total = len(prs.slides)

    for i, slide in enumerate(prs.slides):
        notes = get_slide_notes(slide)
        if not notes:
            print(f"  Slide {i+1}/{total}: no notes, skipping")
            continue

        audio_file = os.path.join(AUDIO_DIR, f"slide_{i+1:02d}.mp3")

        # Reuse cached audio if already downloaded
        if os.path.exists(audio_file):
            print(f"  Slide {i+1}/{total}: using cached audio", end=" ")
            with open(audio_file, "rb") as f:
                audio = f.read()
        else:
            print(f"  Slide {i+1}/{total}: generating audio ({len(notes)} chars)...", end=" ", flush=True)
            try:
                audio = generate_audio(notes)
                with open(audio_file, "wb") as f:
                    f.write(audio)
            except Exception as e:
                print(f"ERROR: {e}")
                continue
            time.sleep(0.5)

        print(f"done ({len(audio)//1024} KB)")
        embed_audio(prs, slide, audio, i)

    prs.save(OUTPUT_PPTX)
    print(f"\nSaved: {OUTPUT_PPTX}")


if __name__ == "__main__":
    main()
